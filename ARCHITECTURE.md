# ARCHITECTURE.md — Mapa de dependencias

Referencia rápida para entender qué depende de qué antes de tocar un archivo.

---

## Grafo de dependencias (de bajo a alto nivel)

```
logging_utils.py          ← no depende de nada interno
    ↑
formularios/common.py     ← depende de: logging_utils
    ↑
google_api_requests.py    ← no depende de nada interno
    ↑
google_sheets_client.py   ← depende de: formularios/common, google_api_requests
drive_upload.py           ← depende de: formularios/common, google_api_requests, logging_utils
dictation.py              ← depende de: formularios/common
text_review.py            ← depende de: formularios/common, logging_utils
                               y también importa módulos de formularios (para schemas)
updater.py                ← depende de: formularios/common, version_info
version_info.py           ← no depende de nada interno

formularios/ui_feedback.py   ← solo depende de tkinter (sin imports internos)
formularios/user_messages.py ← sin dependencias internas
formularios/finalize_validation.py ← sin dependencias internas
completion_payloads.py       ← sin dependencias internas (solo stdlib)

    ↑
formularios/<nombre>/<nombre>.py   ← depende de: formularios/common, google_sheets_client,
                                       formularios/finalize_validation, completion_payloads
    ↑
app.py   ← depende de TODOS los anteriores
```

---

## Tabla de dependencias por archivo

### `logging_utils.py`
- **Importa de internos:** nada
- **Usado por:** `formularios/common.py`, `drive_upload.py`, `text_review.py`, `app.py`, todos los módulos de formularios indirectamente

### `version_info.py`
- **Importa de internos:** nada
- **Usado por:** `updater.py`, `app.py`

### `formularios/ui_feedback.py`
- **Importa de internos:** nada (solo `tkinter`)
- **Usado por:** `app.py` (funciones de feedback inline en ventanas)

### `formularios/user_messages.py`
- **Importa de internos:** nada
- **Usado por:** `app.py` → `map_exception_to_user_message()`

### `formularios/finalize_validation.py`
- **Importa de internos:** nada
- **Usado por:** `app.py`, módulos de formularios

### `completion_payloads.py`
- **Importa de internos:** nada
- **Usado por:** `app.py` → construye el payload final antes de escribir en Sheets

### `google_api_requests.py`
- **Importa de internos:** nada
- **Exporta:** `execute_google_request_with_retry`, `execute_google_create_with_confirmation`
- **Usado por:** `google_sheets_client.py`, `drive_upload.py`

### `formularios/common.py`
- **Importa de internos:** `logging_utils`
- **Exporta:** funciones `_supabase_*`, `_get_*_dir`, `_load_env_file`, `_load_json_config`, caché de empresa, queue de escritura offline
- **Usado por:** prácticamente todo el resto del proyecto
- **IMPORTANTE:** Es el único lugar donde viven las sesiones Supabase y el write worker

### `google_sheets_client.py`
- **Importa de internos:** `formularios/common`, `google_api_requests`
- **Responsabilidad:** Autenticación con service account, lectura/escritura de celdas
- **Usado por:** módulos de formularios, `app.py`

### `drive_upload.py`
- **Importa de internos:** `formularios/common`, `google_api_requests`, `logging_utils`
- **Responsabilidad:** Subida de archivos Excel/PDF a Google Drive en hilo de fondo
- **Usado por:** `app.py`

### `dictation.py`
- **Importa de internos:** `formularios/common`
- **Responsabilidad:** Grabación → transcripción vía Supabase Edge Function
- **Usado por:** `app.py`

### `text_review.py`
- **Importa de internos:** `formularios/common`, `logging_utils`, módulos de formularios (para schemas de revisión)
- **Responsabilidad:** Revisión ortográfica vía Edge Function (GPT-4.1-nano)
- **Usado por:** `app.py`

### `updater.py`
- **Importa de internos:** `formularios/common`, `version_info`
- **Responsabilidad:** Comparar versión actual vs. GitHub Releases, descargar installer
- **Usado por:** `app.py`

### Módulos de formularios (`formularios/<nombre>/<nombre>.py`)
- **Importan de internos:** `formularios/common`, `google_sheets_client`, `formularios/finalize_validation`, `completion_payloads`
- **Exportan:** constantes de mapeo a columnas Sheets, funciones de construcción de payload, función `build_sheet_payload` o similar
- **Usado por:** `app.py` (importa el módulo completo)

### `app.py`
- **Importa de internos:** todos los anteriores
- **Responsabilidad:** UI completa (Tkinter), flujo de navegación, autenticación, borradores, queue de Drive
- **No debe ser importado por nadie** — es el punto de entrada

---

## Capas de la arquitectura

```
┌──────────────────────────────────────────────────────────┐
│                        app.py                            │
│   (UI Tkinter, ventanas, navegación, auth, queues)       │
├──────────────┬───────────────┬───────────────────────────┤
│  formularios │  drive_upload │  dictation / text_review  │
│  (lógica de  │  (Drive API   │  (Edge Functions Supabase) │
│   cada form) │   background) │                           │
├──────────────┴───────────────┴───────────────────────────┤
│       formularios/common.py  +  google_sheets_client.py  │
│   (Supabase HTTP, auth, caché, paths, queue offline)     │
├──────────────────────────────────────────────────────────┤
│   google_api_requests.py  |  logging_utils.py            │
│   (retry Google API)      |  (archivos de log)           │
└──────────────────────────────────────────────────────────┘
```

---

## ¿Dónde buscar si falla X?

| Síntoma | Archivos a revisar |
|---|---|
| No escribe en Sheets | `google_sheets_client.py`, `formularios/<form>/<form>.py` (mapeo de columnas) |
| No sube a Drive | `drive_upload.py:_perform_drive_upload_attempt`, logs en `drive.log` |
| Error de auth Supabase | `formularios/common._supabase_auth_password_login`, `supabase.log` |
| Datos de empresa incorrectos | `formularios/common._find_cached_company_row`, `offline_cache.db` |
| Formulario no finaliza | `app._guard_form_finalization`, `completion_payloads`, `formularios/finalize_validation` |
| UI no responde | `app._run_async_ui_task` (revisar si tarea bloqueante está en hilo UI) |
| Error de encoding al arrancar | `app._run_encoding_health_check` (~línea 2876) |
| Dictado no transcribe | `dictation.py`, `labs.log`, Edge Function `dictate-transcribe` |
| Revisión ortográfica falla | `text_review.py`, Edge Function `text-review-orthography` |
| Auto-update falla | `updater.py`, `updater.log` en Desktop |

---

## Archivos que NO se deben tocar sin entender sus efectos globales

| Archivo | Riesgo |
|---|---|
| `formularios/common.py` | Usado por todo. Cambios en `_supabase_*` afectan todos los formularios |
| `app.py` líneas 1–225 (imports y constantes) | Rompe toda la app si falla un import |
| `google_sheets_client.py` (caché del servicio) | `_SHEETS_SERVICE_CACHE` es singleton — threading sensitive |
| `drive_upload.py` (worker loop) | Hilo de fondo persistente — cuidado con locks |
| `completion_payloads.py` (`PAYLOAD_SCHEMA_VERSION`) | Cambiar rompe compatibilidad con registros guardados |
