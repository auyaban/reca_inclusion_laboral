import { serve } from "https://deno.land/std@0.224.0/http/server.ts";
import { CONTRACT_PROMPT, SCHEMA, SUBSECTION_SPECS, SYSTEM_PROMPT } from "./assets.ts";

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY") ?? "";
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY") ?? "";
const TRANSCRIBE_MODEL =
  Deno.env.get("OPENAI_SELECTION_TRANSCRIBE_MODEL") ?? "gpt-4o-mini-transcribe";
const EXTRACT_MODEL = Deno.env.get("OPENAI_SELECTION_EXTRACT_MODEL") ?? "gpt-5.4-mini";
const DEFAULT_LANGUAGE = "es";
const MAX_AUDIO_MB = Number(Deno.env.get("DICTATION_MAX_AUDIO_MB") ?? "25");
const REQUESTS_PER_MINUTE = Number(Deno.env.get("DICTATION_RPM") ?? "12");
const AUDIO_MINUTES_PER_HOUR = Number(Deno.env.get("DICTATION_AUDIO_MIN_PER_HOUR") ?? "60");
const TRANSCRIBE_TIMEOUT_MS = Number(Deno.env.get("OPENAI_SELECTION_TRANSCRIBE_TIMEOUT_MS") ?? "45000");
const EXTRACT_TIMEOUT_MS = Number(Deno.env.get("OPENAI_SELECTION_EXTRACT_TIMEOUT_MS") ?? "70000");

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

function labsLog(level: "INFO" | "WARN" | "ERROR", message: string, extra?: Record<string, unknown>) {
  const payload = extra && Object.keys(extra).length ? ` ${JSON.stringify(extra)}` : "";
  console.log(`[LABS] [${level}] ${message}${payload}`);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
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
  return Math.max(0, sizeBytes / 1920000);
}

function checkRateLimit(
  userKey: string,
  estimatedMinutes: number,
): { ok: true } | { ok: false; reason: string } {
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

function extractOutputText(payload: any): string {
  const direct = typeof payload?.output_text === "string" ? payload.output_text.trim() : "";
  if (direct) return direct;
  const output = Array.isArray(payload?.output) ? payload.output : [];
  const chunks: string[] = [];
  for (const item of output) {
    const content = Array.isArray(item?.content) ? item.content : [];
    for (const part of content) {
      if (part?.type === "output_text" || part?.type === "text") {
        const text = String(part?.text ?? "").trim();
        if (text) chunks.push(text);
      }
    }
  }
  return chunks.join("\n").trim();
}

function validateStructuredPayload(payload: unknown): string[] {
  const errors: string[] = [];
  if (!isRecord(payload)) {
    return ["La salida estructurada no es un objeto JSON."];
  }
  const expectedTop = [
    "schema_version",
    "form_id",
    "section_id",
    "subsection_key",
    "audio_unit",
    "transcription_summary",
    "warnings",
    "candidate",
  ];
  for (const key of expectedTop) {
    if (!(key in payload)) {
      errors.push(`Falta la clave ${key}.`);
    }
  }
  const warnings = payload.warnings;
  if (!Array.isArray(warnings) || warnings.some((item) => typeof item !== "string")) {
    errors.push("warnings debe ser una lista de strings.");
  }
  if (!isRecord(payload.candidate)) {
    errors.push("candidate debe ser un objeto.");
  }
  return errors;
}

function buildInstructions(subsectionKey: string): string {
  const subsection = SUBSECTION_SPECS?.subsections?.[subsectionKey];
  if (!subsection) {
    throw new Error(`Subseccion no soportada: ${subsectionKey}`);
  }
  const examples = Array.isArray(subsection.examples) ? subsection.examples : [];
  const exampleText = examples.map((value: string) => `- ${value}`).join("\n");
  return [
    SYSTEM_PROMPT,
    "",
    `Subseccion objetivo: ${subsectionKey}`,
    `Titulo: ${String(subsection.title ?? "").trim()}`,
    "",
    `Guia para el profesional: ${String(subsection.script ?? "").trim()}`,
    "",
    "Ejemplos esperados:",
    exampleText,
    "",
    `Reglas especificas de la subseccion: ${String(subsection.prompt_fragment ?? "").trim()}`,
    "",
    "Contrato del formulario y opciones permitidas:",
    CONTRACT_PROMPT,
  ].join("\n");
}

async function transcribeAudio(audio: File, language: string): Promise<string> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort("transcription_timeout"), TRANSCRIBE_TIMEOUT_MS);
  const openaiForm = new FormData();
  openaiForm.set("model", TRANSCRIBE_MODEL);
  openaiForm.set("language", language);
  openaiForm.set("response_format", "json");
  openaiForm.set("file", audio, audio.name || "selection-section2.wav");

  try {
    const upstream = await fetch("https://api.openai.com/v1/audio/transcriptions", {
      method: "POST",
      headers: { Authorization: `Bearer ${OPENAI_API_KEY}` },
      body: openaiForm,
      signal: controller.signal,
    });
    if (!upstream.ok) {
      let message = `OpenAI error ${upstream.status}`;
      try {
        const payloadErr = await upstream.json();
        const msg = payloadErr?.error?.message ?? payloadErr?.message;
        if (typeof msg === "string" && msg.trim()) {
          message = msg.trim();
        }
      } catch {
        // ignore parse errors
      }
      throw new Error(message);
    }

    const data = (await upstream.json()) as { text?: string };
    const textRaw = String(data?.text ?? "");
    return textRaw.replace(/\s+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  } catch (err) {
    if ((err as Error)?.name === "AbortError") {
      throw new Error(`Timeout de transcripcion tras ${TRANSCRIBE_TIMEOUT_MS}ms.`);
    }
    throw err;
  } finally {
    clearTimeout(timeoutId);
  }
}

async function extractStructuredJson(
  transcription: string,
  subsectionKey: string,
  candidateIndex: number,
): Promise<Record<string, unknown>> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort("extraction_timeout"), EXTRACT_TIMEOUT_MS);
  const instructions = buildInstructions(subsectionKey);
  const input = [
    `Contexto fijo: form_id=seleccion_incluyente_labs, section_id=section_2, subsection_key=${subsectionKey}, audio_unit=single_candidate, candidate_index=${candidateIndex}.`,
    "",
    "Transcripcion de voz:",
    transcription,
    "",
    "Devuelve solo el JSON estructurado.",
  ].join("\n");

  try {
    const upstream = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: EXTRACT_MODEL,
        instructions,
        input,
        text: {
          format: {
            type: "json_schema",
            name: "selection_section2_extract",
            strict: true,
            schema: SCHEMA,
          },
        },
      }),
      signal: controller.signal,
    });

    if (!upstream.ok) {
      let message = `OpenAI error ${upstream.status}`;
      try {
        const payloadErr = await upstream.json();
        const msg = payloadErr?.error?.message ?? payloadErr?.message;
        if (typeof msg === "string" && msg.trim()) {
          message = msg.trim();
        }
      } catch {
        // ignore parse errors
      }
      throw new Error(message);
    }

    const payload = await upstream.json();
    const outputText = extractOutputText(payload);
    if (!outputText) {
      throw new Error("OpenAI no devolvio JSON estructurado.");
    }

    let structured: unknown;
    try {
      structured = JSON.parse(outputText);
    } catch {
      throw new Error("La salida estructurada no se pudo parsear como JSON.");
    }

    const errors = validateStructuredPayload(structured);
    if (errors.length) {
      throw new Error(errors.join(" | "));
    }

    return structured as Record<string, unknown>;
  } catch (err) {
    if ((err as Error)?.name === "AbortError") {
      throw new Error(`Timeout de extraccion tras ${EXTRACT_TIMEOUT_MS}ms.`);
    }
    throw err;
  } finally {
    clearTimeout(timeoutId);
  }
}

serve(async (req: Request) => {
  labsLog("INFO", "request_start", {
    method: req.method,
    content_type: req.headers.get("content-type") ?? "",
  });
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
    labsLog("WARN", "invalid_auth", {});
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
  labsLog("INFO", "request_formdata_ok", {});

  const formId = String(formData.get("form_id") ?? "").trim();
  const sectionId = String(formData.get("section_id") ?? "").trim();
  const subsectionKey = String(formData.get("subsection_key") ?? "").trim();
  const candidateIndexRaw = String(formData.get("candidate_index") ?? "").trim();
  const language = String(formData.get("language") ?? DEFAULT_LANGUAGE).trim() || DEFAULT_LANGUAGE;
  const audio = formData.get("audio_file");

  if (formId !== "seleccion_incluyente_labs") {
    return json(400, { ok: false, error: { code: "invalid_form_id", message: "form_id invalido." } });
  }
  if (sectionId !== "section_2") {
    return json(400, { ok: false, error: { code: "invalid_section_id", message: "section_id invalido." } });
  }
  if (!SUBSECTION_SPECS?.subsections?.[subsectionKey]) {
    return json(400, {
      ok: false,
      error: { code: "invalid_subsection", message: "subsection_key invalido." },
    });
  }

  const candidateIndex = Number(candidateIndexRaw);
  if (!Number.isInteger(candidateIndex) || candidateIndex < 1) {
    return json(400, {
      ok: false,
      error: { code: "invalid_candidate_index", message: "candidate_index debe ser un entero positivo." },
    });
  }

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

  const rateLimit = checkRateLimit(userKey, estimateAudioMinutesFromWavBytes(audio.size));
  if (!rateLimit.ok) {
    labsLog("WARN", "rate_limited", {
      subsection_key: subsectionKey,
      candidate_index: candidateIndex,
    });
    return json(429, { ok: false, error: { code: "rate_limited", message: rateLimit.reason } });
  }

  labsLog("INFO", "request_received", {
    subsection_key: subsectionKey,
    candidate_index: candidateIndex,
    language,
    audio_size: audio.size,
    transcribe_model: TRANSCRIBE_MODEL,
    extract_model: EXTRACT_MODEL,
  });

  let transcription = "";
  try {
    transcription = await transcribeAudio(audio, language);
  } catch (err) {
    labsLog("ERROR", "transcription_failed", {
      subsection_key: subsectionKey,
      candidate_index: candidateIndex,
      detail: String((err as Error)?.message || "Error de transcripcion."),
    });
    return json(502, {
      ok: false,
      error: { code: "openai_transcription_error", message: String((err as Error)?.message || "Error de transcripcion.") },
    });
  }

  if (!transcription) {
    labsLog("WARN", "empty_transcription", {
      subsection_key: subsectionKey,
      candidate_index: candidateIndex,
    });
    return json(422, {
      ok: false,
      error: { code: "empty_transcription", message: "OpenAI no devolvio texto." },
    });
  }

  labsLog("INFO", "transcription_ok", {
    subsection_key: subsectionKey,
    candidate_index: candidateIndex,
    transcription_chars: transcription.length,
  });

  let extraction: Record<string, unknown>;
  try {
    extraction = await extractStructuredJson(transcription, subsectionKey, candidateIndex);
  } catch (err) {
    labsLog("ERROR", "extraction_failed", {
      subsection_key: subsectionKey,
      candidate_index: candidateIndex,
      detail: String((err as Error)?.message || "Error de extraccion."),
    });
    return json(502, {
      ok: false,
      error: { code: "openai_extraction_error", message: String((err as Error)?.message || "Error de extraccion.") },
      transcription,
    });
  }

  const warnings = Array.isArray(extraction.warnings)
    ? extraction.warnings.filter((item) => typeof item === "string")
    : [];

  labsLog("INFO", "extraction_ok", {
    subsection_key: subsectionKey,
    candidate_index: candidateIndex,
    warnings: warnings.length,
    candidate_fields: isRecord(extraction.candidate) ? Object.keys(extraction.candidate).length : 0,
  });

  return json(200, {
    ok: true,
    transcription,
    extraction,
    warnings,
    usage: {
      transcribe_model: TRANSCRIBE_MODEL,
      extract_model: EXTRACT_MODEL,
    },
    debug: {
      subsection_key: subsectionKey,
      candidate_index: candidateIndex,
    },
  });
});
