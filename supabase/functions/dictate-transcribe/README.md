# Dictate Transcribe (Supabase Edge Function)

Esta función recibe audio (`wav`) desde la app de escritorio, valida JWT y envía la transcripción a OpenAI.

## Requisitos

1. Supabase CLI instalado y autenticado.
2. Proyecto Supabase linkeado.
3. Variables secretas configuradas en Supabase:
   - `OPENAI_API_KEY` (obligatoria)
   - `OPENAI_STT_MODEL` (opcional, default `gpt-4o-mini-transcribe`)
   - `OPENAI_STT_LANGUAGE` (opcional, default `es`)
   - `DICTATION_MAX_AUDIO_MB` (opcional, default `25`)
   - `DICTATION_RPM` (opcional, default `12`)
   - `DICTATION_AUDIO_MIN_PER_HOUR` (opcional, default `60`)

## Comandos

```bash
supabase functions deploy dictate-transcribe --project-ref <PROJECT_REF> --verify-jwt
```

```bash
supabase secrets set OPENAI_API_KEY=sk-... --project-ref <PROJECT_REF>
```

Opcional:

```bash
supabase secrets set OPENAI_STT_MODEL=gpt-4o-mini-transcribe --project-ref <PROJECT_REF>
supabase secrets set OPENAI_STT_LANGUAGE=es --project-ref <PROJECT_REF>
supabase secrets set DICTATION_MAX_AUDIO_MB=25 --project-ref <PROJECT_REF>
supabase secrets set DICTATION_RPM=12 --project-ref <PROJECT_REF>
supabase secrets set DICTATION_AUDIO_MIN_PER_HOUR=60 --project-ref <PROJECT_REF>
```

## Endpoint esperado por la app

`POST {SUPABASE_URL}/functions/v1/dictate-transcribe`

Headers:
- `apikey: <SUPABASE_KEY>`
- `Authorization: Bearer <access_token_usuario>`

Body (`multipart/form-data`):
- `audio_file`: archivo WAV
- `form_id`: string
- `field_id`: string
- `language`: string (ej. `es`)

