import { CONTRACT_PROMPT, SCHEMA, SUBSECTION_SPECS, SYSTEM_PROMPT } from "./assets.ts";

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY") ?? "";
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY") ?? "";
const TRANSCRIBE_MODEL =
  Deno.env.get("OPENAI_VACANCY_TRANSCRIBE_MODEL") ?? "gpt-4o-mini-transcribe";
const EXTRACT_MODEL = Deno.env.get("OPENAI_VACANCY_EXTRACT_MODEL") ?? "gpt-5.4-mini";
const DEFAULT_LANGUAGE = "es";
const MAX_AUDIO_MB = Number(Deno.env.get("DICTATION_MAX_AUDIO_MB") ?? "25");
const REQUESTS_PER_MINUTE = Number(Deno.env.get("DICTATION_RPM") ?? "12");
const AUDIO_MINUTES_PER_HOUR = Number(Deno.env.get("DICTATION_AUDIO_MIN_PER_HOUR") ?? "60");
const TRANSCRIBE_TIMEOUT_MS = Number(Deno.env.get("OPENAI_VACANCY_TRANSCRIBE_TIMEOUT_MS") ?? "45000");
const EXTRACT_TIMEOUT_MS = Number(Deno.env.get("OPENAI_VACANCY_EXTRACT_TIMEOUT_MS") ?? "70000");

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

function log(level: "INFO" | "WARN" | "ERROR", message: string, extra?: Record<string, unknown>) {
  const suffix = extra && Object.keys(extra).length ? ` ${JSON.stringify(extra)}` : "";
  console.log(`[VACANCY_VOICE] [${level}] ${message}${suffix}`);
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
    "semantic",
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
  if (!isRecord(payload.semantic)) {
    errors.push("semantic debe ser un objeto.");
  }
  if (!isRecord(payload.candidate)) {
    errors.push("candidate debe ser un objeto.");
  }
  return errors;
}

function buildInstructions(subsectionKey: string): string {
  const subsection = SUBSECTION_SPECS?.subsections?.[subsectionKey as keyof typeof SUBSECTION_SPECS.subsections];
  if (!subsection) {
    throw new Error(`Subseccion no soportada: ${subsectionKey}`);
  }
  const questions = Array.isArray(subsection.questions) ? subsection.questions : [];
  const examples = Array.isArray(subsection.examples) ? subsection.examples : [];
  return [
    SYSTEM_PROMPT,
    "",
    `Subseccion objetivo: ${subsectionKey}`,
    `Titulo: ${String(subsection.title ?? "").trim()}`,
    "",
    `Guia para el profesional: ${String(subsection.script ?? "").trim()}`,
    "",
    "Preguntas que el audio deberia responder:",
    questions.map((value: string) => `- ${value}`).join("\n"),
    "",
    "Ejemplo breve:",
    examples.map((value: string) => `- ${value}`).join("\n"),
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
  openaiForm.set("file", audio, audio.name || "vacancy-section2.wav");

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
  sectionId: string,
  subsectionKey: string,
): Promise<Record<string, unknown>> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort("extraction_timeout"), EXTRACT_TIMEOUT_MS);
  const instructions = buildInstructions(subsectionKey);
  const input = [
    `Contexto fijo: form_id=condiciones_vacante, section_id=${sectionId}, subsection_key=${subsectionKey}, audio_unit=single_section.`,
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
            name: "vacancy_section2_extract",
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
    const text = extractOutputText(payload);
    if (!text) {
      throw new Error("La respuesta del modelo llego vacia.");
    }
    return JSON.parse(text) as Record<string, unknown>;
  } catch (err) {
    if ((err as Error)?.name === "AbortError") {
      throw new Error(`Timeout de extraccion tras ${EXTRACT_TIMEOUT_MS}ms.`);
    }
    throw err;
  } finally {
    clearTimeout(timeoutId);
  }
}

Deno.serve(async (request) => {
  if (request.method !== "POST") {
    return json(405, { ok: false, error: { code: "method_not_allowed", message: "Metodo no permitido." } });
  }
  if (!OPENAI_API_KEY) {
    return json(500, {
      ok: false,
      error: { code: "missing_openai_key", message: "OPENAI_API_KEY no configurada." },
    });
  }

  const auth = request.headers.get("Authorization") ?? "";
  const jwt = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";
  if (!jwt) {
    return json(401, { ok: false, error: { code: "missing_auth", message: "JWT requerido." } });
  }

  let userId = "";
  try {
    userId = await resolveUserId(jwt);
  } catch (err) {
    return json(401, {
      ok: false,
      error: { code: "invalid_auth", message: (err as Error).message },
    });
  }

  let form: FormData;
  try {
    form = await request.formData();
  } catch {
    return json(400, {
      ok: false,
      error: { code: "invalid_form", message: "No se pudo leer multipart/form-data." },
    });
  }

  const audio = form.get("audio_file");
  const formId = String(form.get("form_id") ?? "").trim();
  const sectionId = String(form.get("section_id") ?? "").trim();
  const subsectionKey = String(form.get("subsection_key") ?? "").trim();
  if (!(audio instanceof File)) {
    return json(400, { ok: false, error: { code: "missing_audio", message: "Falta audio_file." } });
  }
  if (formId !== "condiciones_vacante") {
    return json(400, { ok: false, error: { code: "invalid_form_id", message: "form_id invalido." } });
  }
  const validSectionId =
    (subsectionKey === "section_2_vacancy" && sectionId === "section_2")
    || (subsectionKey === "section_2_1_schedule_experience" && sectionId === "section_2_1");
  if (!validSectionId) {
    return json(400, {
      ok: false,
      error: { code: "invalid_section_or_subsection", message: "section_id o subsection_key invalido." },
    });
  }

  if (audio.size <= 0) {
    return json(400, { ok: false, error: { code: "empty_audio", message: "El audio esta vacio." } });
  }
  if (audio.size > MAX_AUDIO_MB * 1024 * 1024) {
    return json(413, {
      ok: false,
      error: { code: "audio_too_large", message: `El audio supera ${MAX_AUDIO_MB}MB.` },
    });
  }

  const estimatedMinutes = estimateAudioMinutesFromWavBytes(audio.size);
  const rate = checkRateLimit(userId, estimatedMinutes);
  if (!rate.ok) {
    return json(429, { ok: false, error: { code: "rate_limited", message: rate.reason } });
  }

  const language = String(form.get("language") ?? DEFAULT_LANGUAGE).trim() || DEFAULT_LANGUAGE;
  const started = Date.now();

  try {
    const transcription = await transcribeAudio(audio, language);
    if (!transcription) {
      return json(422, {
        ok: false,
        error: { code: "empty_transcription", message: "La transcripcion llego vacia." },
      });
    }

    const extraction = await extractStructuredJson(transcription, sectionId, subsectionKey);
    const validationErrors = validateStructuredPayload(extraction);
    if (validationErrors.length) {
      log("ERROR", "structured_payload_invalid", { validationErrors });
      return json(502, {
        ok: false,
        error: {
          code: "invalid_structured_payload",
          message: `La respuesta estructurada no es valida: ${validationErrors.join(" | ")}`,
        },
        transcription,
      });
    }

    const usage = {
      transcribe_model: TRANSCRIBE_MODEL,
      extract_model: EXTRACT_MODEL,
      elapsed_ms: Date.now() - started,
    };
    log("INFO", "voice_submit_ok", { elapsed_ms: usage.elapsed_ms, userId });
    return json(200, {
      ok: true,
      transcription,
      extraction,
      warnings: extraction.warnings ?? [],
      usage,
      debug: {
        subsection_key: subsectionKey,
        section_id: sectionId,
      },
    });
  } catch (err) {
    const message = (err as Error)?.message || "No fue posible procesar el audio.";
    log("ERROR", "voice_submit_failed", { message, userId });
    return json(502, {
      ok: false,
      error: { code: "openai_processing_error", message },
    });
  }
});
