# selection-section2-extract

Edge Function protegida para el experimento de voz en `Seleccion Incluyente Labs -> section_2`.

Flujo:
- Valida JWT del usuario con Supabase Auth.
- Recibe `multipart/form-data` con un audio corto y metadatos de subseccion.
- Transcribe con OpenAI Speech-to-Text.
- Extrae JSON estructurado con OpenAI Responses API + `json_schema`.
- Devuelve transcripcion, payload estructurado y advertencias.

Campos `multipart/form-data` requeridos:
- `form_id=seleccion_incluyente_labs`
- `section_id=section_2`
- `subsection_key`
- `candidate_index`
- `language`
- `audio_file`

Variables de entorno:
- `OPENAI_API_KEY`
- `SUPABASE_URL`
- `SUPABASE_ANON_KEY`
- `OPENAI_SELECTION_TRANSCRIBE_MODEL`
- `OPENAI_SELECTION_EXTRACT_MODEL`

Defaults:
- `OPENAI_SELECTION_TRANSCRIBE_MODEL=gpt-4o-mini-transcribe`
- `OPENAI_SELECTION_EXTRACT_MODEL=gpt-5.4-mini`

Deploy:
```bash
supabase functions deploy selection-section2-extract --project-ref <PROJECT_REF> --verify-jwt
```
