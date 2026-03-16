import { serve } from "https://deno.land/std@0.224.0/http/server.ts";

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY") ?? "";
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY") ?? "";
const OPENAI_MODEL = Deno.env.get("OPENAI_TEXT_REVIEW_MODEL") ?? "gpt-4.1-mini";
const MAX_TEXT_CHARS = Number(Deno.env.get("OPENAI_TEXT_REVIEW_MAX_CHARS") ?? "6000");

const REVIEW_PROMPT =
  "Corrige solo ortografia, tildes, signos de puntuacion y uso basico de mayusculas/minusculas. " +
  "No resumas, no reformules, no cambies el tono, no inventes informacion y no alteres el sentido del texto. " +
  "No cambies nombres propios, numeros, correos, URLs, siglas, articulos legales, referencias normativas, codigos, " +
  "ni el formato general de listas o parrafos. Devuelve unicamente el texto final corregido en texto plano.";

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
  try {
    await resolveUserId(jwt);
  } catch (err) {
    return json(401, {
      ok: false,
      error: { code: "invalid_auth", message: String((err as Error)?.message || "JWT invalido.") },
    });
  }

  let body: any;
  try {
    body = await req.json();
  } catch {
    return json(400, { ok: false, error: { code: "invalid_json", message: "JSON invalido." } });
  }

  const text = String(body?.text ?? "").trim();
  const model = String(body?.model ?? OPENAI_MODEL).trim() || OPENAI_MODEL;
  if (!text) {
    return json(400, { ok: false, error: { code: "missing_text", message: "text es obligatorio." } });
  }
  if (text.length > MAX_TEXT_CHARS) {
    return json(413, {
      ok: false,
      error: { code: "text_too_large", message: `El texto supera ${MAX_TEXT_CHARS} caracteres.` },
    });
  }

  let upstream: Response;
  try {
    upstream = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        instructions: REVIEW_PROMPT,
        input: text,
      }),
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

  let payload: any;
  try {
    payload = await upstream.json();
  } catch {
    return json(502, {
      ok: false,
      error: { code: "invalid_openai_response", message: "Respuesta de OpenAI invalida." },
    });
  }

  const reviewedText = extractOutputText(payload);
  if (!reviewedText) {
    return json(422, {
      ok: false,
      error: { code: "empty_review", message: "OpenAI no devolvio texto corregido." },
    });
  }

  return json(200, {
    ok: true,
    text: reviewedText,
    usage: {
      model,
    },
  });
});
