import { serve } from "https://deno.land/std@0.224.0/http/server.ts";

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY") ?? "";
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY") ?? "";
const OPENAI_MODEL = Deno.env.get("OPENAI_STT_MODEL") ?? "gpt-4o-mini-transcribe";
const DEFAULT_LANGUAGE = Deno.env.get("OPENAI_STT_LANGUAGE") ?? "es";
const MAX_AUDIO_MB = Number(Deno.env.get("DICTATION_MAX_AUDIO_MB") ?? "25");
const REQUESTS_PER_MINUTE = Number(Deno.env.get("DICTATION_RPM") ?? "12");
const AUDIO_MINUTES_PER_HOUR = Number(Deno.env.get("DICTATION_AUDIO_MIN_PER_HOUR") ?? "60");

type RateState = {
  reqWindowStart: number;
  reqCount: number;
  minWindowStart: number;
  audioMinutes: number;
};

const rateByUser = new Map<string, RateState>();

function json(status: number, payload: unknown) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });
}

async function resolveUserId(jwt: string): Promise<string> {
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
    throw new Error("Supabase auth env no configurado.");
  }
  const resp = await fetch(`${SUPABASE_URL.replace(/\/+$/, "")}/auth/v1/user`, {
    method: "GET",
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${jwt}`,
    },
  });
  if (!resp.ok) {
    throw new Error(`JWT invalido (${resp.status})`);
  }
  const data = (await resp.json()) as { id?: string };
  const userId = String(data?.id ?? "").trim();
  if (!userId) {
    throw new Error("No se pudo resolver usuario desde JWT.");
  }
  return userId;
}

function estimateAudioMinutesFromWavBytes(sizeBytes: number): number {
  // WAV PCM16 mono 16kHz ~= 1,920,000 bytes/min + small header.
  return Math.max(0, sizeBytes / 1920000);
}

function checkRateLimit(userKey: string, estimatedMinutes: number): { ok: true } | { ok: false; reason: string } {
  const now = Date.now();
  const state = rateByUser.get(userKey) ?? {
    reqWindowStart: now,
    reqCount: 0,
    minWindowStart: now,
    audioMinutes: 0,
  };

  if (now - state.reqWindowStart >= 60_000) {
    state.reqWindowStart = now;
    state.reqCount = 0;
  }
  if (now - state.minWindowStart >= 3_600_000) {
    state.minWindowStart = now;
    state.audioMinutes = 0;
  }

  state.reqCount += 1;
  state.audioMinutes += estimatedMinutes;
  rateByUser.set(userKey, state);

  if (state.reqCount > REQUESTS_PER_MINUTE) {
    return { ok: false, reason: "Rate limit por minuto excedido." };
  }
  if (state.audioMinutes > AUDIO_MINUTES_PER_HOUR) {
    return { ok: false, reason: "Rate limit de minutos por hora excedido." };
  }
  return { ok: true };
}

serve(async (req: Request) => {
  if (req.method !== "POST") {
    return json(405, { ok: false, error: { code: "method_not_allowed", message: "Use POST." } });
  }
  if (!OPENAI_API_KEY) {
    return json(500, {
      ok: false,
      error: { code: "missing_openai_key", message: "OPENAI_API_KEY no configurada." },
    });
  }

  const authHeader = req.headers.get("Authorization") ?? "";
  if (!authHeader.toLowerCase().startsWith("bearer ")) {
    return json(401, { ok: false, error: { code: "missing_auth", message: "JWT requerido." } });
  }
  const jwt = authHeader.slice(7).trim();
  let userKey = "";
  try {
    userKey = await resolveUserId(jwt);
  } catch (err) {
    return json(401, {
      ok: false,
      error: { code: "invalid_auth", message: String((err as Error)?.message || "JWT invalido.") },
    });
  }

  const contentType = req.headers.get("content-type") ?? "";
  if (!contentType.toLowerCase().includes("multipart/form-data")) {
    return json(400, {
      ok: false,
      error: { code: "invalid_content_type", message: "Se requiere multipart/form-data." },
    });
  }

  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return json(400, { ok: false, error: { code: "invalid_form", message: "Formulario invalido." } });
  }

  const audio = formData.get("audio_file");
  if (!(audio instanceof File)) {
    return json(400, { ok: false, error: { code: "missing_audio", message: "audio_file es obligatorio." } });
  }
  if (!audio.size) {
    return json(400, { ok: false, error: { code: "empty_audio", message: "El audio esta vacio." } });
  }

  const maxBytes = MAX_AUDIO_MB * 1024 * 1024;
  if (audio.size > maxBytes) {
    return json(413, {
      ok: false,
      error: { code: "audio_too_large", message: `Audio supera ${MAX_AUDIO_MB}MB.` },
    });
  }

  const estimatedMinutes = estimateAudioMinutesFromWavBytes(audio.size);
  const rl = checkRateLimit(userKey, estimatedMinutes);
  if (!rl.ok) {
    return json(429, { ok: false, error: { code: "rate_limited", message: rl.reason } });
  }

  const language = String(formData.get("language") ?? DEFAULT_LANGUAGE).trim() || DEFAULT_LANGUAGE;

  const openaiForm = new FormData();
  openaiForm.set("model", OPENAI_MODEL);
  openaiForm.set("language", language);
  openaiForm.set("response_format", "json");
  openaiForm.set("file", audio, audio.name || "dictation.wav");

  let upstream: Response;
  try {
    upstream = await fetch("https://api.openai.com/v1/audio/transcriptions", {
      method: "POST",
      headers: { Authorization: `Bearer ${OPENAI_API_KEY}` },
      body: openaiForm,
    });
  } catch {
    return json(502, {
      ok: false,
      error: { code: "openai_unreachable", message: "No fue posible conectar con OpenAI." },
    });
  }

  if (!upstream.ok) {
    let message = `OpenAI error ${upstream.status}`;
    try {
      const payloadErr = await upstream.json();
      const msg = payloadErr?.error?.message ?? payloadErr?.message;
      if (typeof msg === "string" && msg.trim()) message = msg.trim();
    } catch {
      // ignore parse errors
    }
    return json(502, { ok: false, error: { code: "openai_error", message } });
  }

  let data: { text?: string };
  try {
    data = await upstream.json();
  } catch {
    return json(502, {
      ok: false,
      error: { code: "invalid_openai_response", message: "Respuesta de OpenAI invalida." },
    });
  }

  const textRaw = String(data?.text ?? "");
  const text = textRaw.replace(/\s+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();

  if (!text) {
    return json(422, {
      ok: false,
      error: { code: "empty_transcription", message: "OpenAI no devolvio texto." },
    });
  }

  return json(200, {
    ok: true,
    text,
    usage: {
      duration_seconds: null,
      model: OPENAI_MODEL,
    },
  });
});
