# Integracion De Dictado (OpenAI + Supabase Edge)

## 1) Dependencias cliente (app desktop)

Instalar en entorno local:

```bash
pip install -r requirements.txt
```

Dependencias nuevas:
- `sounddevice`
- `numpy`

## 2) Desplegar Edge Function

Código: `supabase/functions/dictate-transcribe/index.ts`

```bash
supabase functions deploy dictate-transcribe --project-ref <PROJECT_REF> --verify-jwt
```

## 3) Configurar secrets en Supabase

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

## 4) Variables opcionales en app (.env)

Estas son opcionales; si no existen se usan defaults.

```env
DICTATION_FUNCTION_NAME=dictate-transcribe
DICTATION_LANGUAGE=es
```

## 5) Flujo funcional

1. Usuario pulsa `Dictar` en un campo largo (`tk.Text`).
2. App graba audio temporal local.
3. Usuario pulsa `Detener`.
4. App envia audio a Edge Function con JWT de sesión.
5. Edge Function transcribe con OpenAI.
6. App inserta texto en el cursor y borra audio temporal si fue exitoso.
7. App limpia audios viejos (>24h) al iniciar.

