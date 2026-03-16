# Text Review Orthography (Supabase Edge Function)

Corrige ortografía final de textos libres antes de exportar Excel, usando OpenAI del lado servidor.

## Deploy

```bash
supabase functions deploy text-review-orthography --project-ref <PROJECT_REF> --verify-jwt
```

## Secrets requeridos

```bash
supabase secrets set OPENAI_API_KEY=sk-... --project-ref <PROJECT_REF>
```

Opcional:

```bash
supabase secrets set OPENAI_TEXT_REVIEW_MODEL=gpt-4.1-mini --project-ref <PROJECT_REF>
supabase secrets set OPENAI_TEXT_REVIEW_MAX_CHARS=6000 --project-ref <PROJECT_REF>
```

## Variables opcionales en la app

Si quieres renombrar la función:

```env
OPENAI_TEXT_REVIEW_FUNCTION_NAME=text-review-orthography
```

## Request

`POST {SUPABASE_URL}/functions/v1/text-review-orthography`

Headers:
- `apikey: <SUPABASE_KEY>`
- `Authorization: Bearer <access_token_usuario>`
- `Content-Type: application/json`

Body:

```json
{
  "text": "Texto a corregir",
  "model": "gpt-4.1-mini"
}
```

## Response

```json
{
  "ok": true,
  "text": "Texto corregido",
  "usage": {
    "model": "gpt-4.1-mini"
  }
}
```
