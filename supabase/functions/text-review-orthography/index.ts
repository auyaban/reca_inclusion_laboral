import { serve } from "https://deno.land/std@0.224.0/http/server.ts";

const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY") ?? "";
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY") ?? "";
const OPENAI_MODEL = Deno.env.get("OPENAI_TEXT_REVIEW_MODEL") ?? "gpt-4.1-nano";
const MAX_TEXT_CHARS = Number(Deno.env.get("OPENAI_TEXT_REVIEW_MAX_CHARS") ?? "6000");
const MAX_BATCH_ITEMS = Number(Deno.env.get("OPENAI_TEXT_REVIEW_BATCH_MAX_ITEMS") ?? "8");
const MAX_BATCH_CHARS = Number(Deno.env.get("OPENAI_TEXT_REVIEW_BATCH_MAX_CHARS") ?? "12000");
const LIST_FORMATTING_PROMPT_SUFFIX =
  "Si el texto ya representa una enumeracion evidente, por ejemplo marcadores numerados, " +
  "items en lineas separadas, una frase introductoria terminada en dos puntos seguida de varias lineas cortas, " +
  "o elementos claramente separados por punto y coma, puedes " +
  "devolverlo como lista simple en texto plano usando prefijos '- '. " +
  "Hazlo solo cuando la estructura enumerativa sea inequívoca y no haya riesgo de convertir " +
  "un parrafo normal en lista. Si no es inequívoco, conserva el formato original.";

const REVIEW_PROMPT =
  "Corrige solo ortografia, tildes, signos de puntuacion y uso basico de mayusculas/minusculas. " +
  "No resumas, no reformules, no cambies el tono, no inventes informacion y no alteres el sentido del texto. " +
  "No cambies nombres propios, numeros, correos, URLs, siglas, articulos legales, referencias normativas, codigos, " +
  "ni el formato general de listas o parrafos. " +
  `${LIST_FORMATTING_PROMPT_SUFFIX} ` +
  "Devuelve unicamente el texto final corregido en texto plano.";

const BATCH_REVIEW_PROMPT =
  "Corrige solo ortografia, tildes, signos de puntuacion y uso basico de mayusculas/minusculas. " +
  'Recibiras un JSON con este formato exacto: {"items":[{"id":"...","text":"..."}]}. ' +
  "Corrige cada campo text por separado, sin mezclar items entre si. " +
  "No resumas, no reformules, no cambies el tono, no inventes informacion y no alteres el sentido del texto. " +
  "No cambies nombres propios, numeros, correos, URLs, siglas, articulos legales, referencias normativas, codigos, " +
  "ni el formato general de listas o parrafos. " +
  `${LIST_FORMATTING_PROMPT_SUFFIX} ` +
  'Devuelve exclusivamente JSON valido con este mismo formato: {"items":[{"id":"...","text":"texto corregido"}]}. ' +
  "Manten exactamente los mismos ids y la misma cantidad de items. No agregues markdown ni explicaciones.";

type ReviewItem = {
  id: string;
  text: string;
};

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

function stripCodeFences(text: string): string {
  const value = String(text ?? "").trim();
  if (!value.startsWith("```")) return value;
  const lines = value.split(/\r?\n/);
  if (lines.length && lines[0].trim().startsWith("```")) lines.shift();
  if (lines.length && lines[lines.length - 1].trim().startsWith("```")) lines.pop();
  return lines.join("\n").trim();
}

function extractJsonCandidate(text: string): string {
  const value = stripCodeFences(text);
  if (value.startsWith("{") && value.endsWith("}")) return value;
  const start = value.indexOf("{");
  const end = value.lastIndexOf("}");
  if (start !== -1 && end > start) return value.slice(start, end + 1);
  return value;
}

function parseBatchResult(text: string, expectedIds: string[]): ReviewItem[] {
  const parsed = JSON.parse(extractJsonCandidate(text));
  const items = Array.isArray(parsed?.items) ? parsed.items : null;
  if (!items) {
    throw new Error("La respuesta por lotes no contiene items validos.");
  }
  const expectedSet = new Set(expectedIds);
  const reviewedMap = new Map<string, string>();
  for (const item of items) {
    const itemId = String(item?.id ?? "").trim();
    if (!itemId || !expectedSet.has(itemId)) {
      throw new Error("La respuesta por lotes devolvio ids inesperados.");
    }
    if (reviewedMap.has(itemId)) {
      throw new Error("La respuesta por lotes devolvio ids duplicados.");
    }
    reviewedMap.set(itemId, String(item?.text ?? "").trim());
  }
  for (const itemId of expectedIds) {
    if (!reviewedMap.has(itemId)) {
      throw new Error("La respuesta por lotes no devolvio todos los items.");
    }
  }
  return expectedIds.map((id) => ({ id, text: reviewedMap.get(id) ?? "" }));
}

async function callOpenAI(input: string, model: string, instructions: string): Promise<string> {
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
        instructions,
        input,
      }),
    });
  } catch {
    throw new Error("No fue posible conectar con OpenAI.");
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
    throw new Error(message);
  }

  const payload = await upstream.json();
  const reviewedText = extractOutputText(payload);
  if (!reviewedText) {
    throw new Error("OpenAI no devolvio texto corregido.");
  }
  return reviewedText;
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

  const model = String(body?.model ?? OPENAI_MODEL).trim() || OPENAI_MODEL;
  const rawItems = Array.isArray(body?.items) ? body.items : [];
  const text = String(body?.text ?? "").trim();

  if (rawItems.length > 0) {
    if (rawItems.length > MAX_BATCH_ITEMS) {
      return json(413, {
        ok: false,
        error: {
          code: "batch_too_large",
          message: `El lote supera ${MAX_BATCH_ITEMS} items.`,
        },
      });
    }
    const items: ReviewItem[] = [];
    const seenIds = new Set<string>();
    let totalChars = 0;
    for (const rawItem of rawItems) {
      const itemId = String(rawItem?.id ?? "").trim();
      const itemText = String(rawItem?.text ?? "").trim();
      if (!itemId || !itemText) {
        return json(400, {
          ok: false,
          error: { code: "invalid_batch_item", message: "Cada item debe incluir id y text." },
        });
      }
      if (seenIds.has(itemId)) {
        return json(400, {
          ok: false,
          error: { code: "duplicate_batch_item", message: "Los ids del lote deben ser unicos." },
        });
      }
      if (itemText.length > MAX_TEXT_CHARS) {
        return json(413, {
          ok: false,
          error: {
            code: "text_too_large",
            message: `Un item supera ${MAX_TEXT_CHARS} caracteres.`,
          },
        });
      }
      seenIds.add(itemId);
      totalChars += itemText.length;
      items.push({ id: itemId, text: itemText });
    }
    if (totalChars > MAX_BATCH_CHARS) {
      return json(413, {
        ok: false,
        error: {
          code: "batch_chars_too_large",
          message: `El lote supera ${MAX_BATCH_CHARS} caracteres.`,
        },
      });
    }

    try {
      const reviewedText = await callOpenAI(
        JSON.stringify({ items }),
        model,
        BATCH_REVIEW_PROMPT,
      );
      const reviewedItems = parseBatchResult(
        reviewedText,
        items.map((item) => item.id),
      );
      return json(200, {
        ok: true,
        items: reviewedItems,
        usage: {
          model,
        },
      });
    } catch (err) {
      return json(502, {
        ok: false,
        error: {
          code: "openai_error",
          message: String((err as Error)?.message || "OpenAI no respondio correctamente."),
        },
      });
    }
  }

  if (!text) {
    return json(400, { ok: false, error: { code: "missing_text", message: "text es obligatorio." } });
  }
  if (text.length > MAX_TEXT_CHARS) {
    return json(413, {
      ok: false,
      error: { code: "text_too_large", message: `El texto supera ${MAX_TEXT_CHARS} caracteres.` },
    });
  }

  try {
    const reviewedText = await callOpenAI(text, model, REVIEW_PROMPT);
    return json(200, {
      ok: true,
      text: reviewedText,
      usage: {
        model,
      },
    });
  } catch (err) {
    return json(502, {
      ok: false,
      error: {
        code: "openai_error",
        message: String((err as Error)?.message || "OpenAI no respondio correctamente."),
      },
    });
  }
});
