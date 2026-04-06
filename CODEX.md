# CODEX.md - Guia operativa para Codex

Este archivo resume el contexto minimo que necesito para trabajar en este repo sin releer toda la documentacion cada vez.
Fuente principal: `ARCHITECTURE.md`, `CLAUDE.md`, `CHANGELOG.md`, `README_INSTALL.md`, `docs/*.md`, `supabase/functions/*/README.md`, `.env.example`, `requirements.txt`.

## 1. Que es este proyecto

Aplicacion de escritorio Windows en Python + Tkinter para diligenciar formularios de inclusion laboral.
La app:

- autentica usuarios contra Supabase
- consulta y cachea datos de empresa
- autosalva formularios y usa cola offline para escrituras
- exporta informacion a Google Sheets / Excel
- sube archivos a Google Drive
- soporta dictado y revision ortografica via Supabase Edge Functions
- se distribuye como ejecutable PyInstaller con instalador Inno Setup

## 2. Mapa rapido de arquitectura

Capas principales:

1. `app.py`
   Punto de entrada, UI completa Tkinter, navegacion, auth, drafts, finalizacion, colas y ventanas.
2. Formularios `formularios/<nombre>/<nombre>.py`
   Logica por formulario, mappings a Sheets, payloads, validacion y cache local por modulo.
3. Servicios compartidos
   - `formularios/common.py`: Supabase HTTP, rutas, cache, colas offline, sesiones, utilidades globales.
   - `google_sheets_client.py`: lectura/escritura Sheets y service account.
   - `drive_upload.py`: worker de subida a Drive.
   - `dictation.py`: audio -> Edge Function -> transcripcion.
   - `text_review.py`: texto -> Edge Function -> correccion ortografica.
   - `updater.py`: auto-update via GitHub Releases.
4. Base tecnica
   - `google_api_requests.py`: retries/backoff Google API.
   - `logging_utils.py`: logs por canal.
   - `version_info.py`: version y rutas de datos.

Dependencia critica:

- `formularios/common.py` es el centro del runtime compartido. Cambios aqui pueden afectar login, cache, colas, autosave y escritura offline en toda la app.

## 3. Archivos que debo ubicar primero

- `app.py`: archivo dominante del proyecto. No asumir nada sin ubicar la ventana o helper correcto.
- `formularios/common.py`: auth Supabase, cache, rutas, write queue.
- `google_sheets_client.py`: bugs de escritura/lectura Sheets.
- `drive_upload.py`: bugs de Drive y cola de archivos.
- `dictation.py`: integracion de voz.
- `text_review.py`: integracion de correccion ortografica.
- `updater.py`: releases y actualizacion en cliente.
- `completion_payloads.py`: payload final antes de escribir/exportar.
- `formularios/finalize_validation.py`: validacion antes de finalizar.
- `formularios/ui_feedback.py`: feedback visual de campos.
- `formularios/user_messages.py`: mensajes amigables al usuario.

## 4. Formularios y ventanas

Formularios productivos principales:

- `presentacion_programa`
- `evaluacion_programa`
- `condiciones_vacante`
- `seleccion_incluyente`
- `contratacion_incluyente`
- `induccion_organizacional`
- `induccion_operativa`
- `sensibilizacion`
- `seguimientos`

Ventanas importantes en `app.py`:

- `HubWindow`: menu principal
- `Section1Window`: busqueda y confirmacion de empresa
- `EvaluacionAccesibilidadWindow`
- `CondicionesVacanteWindow`
- `SeleccionIncluyenteWindow`
- `ContratacionIncluyenteWindow`
- `InduccionOrganizacionalWindow`
- `InduccionOperativaWindow`
- `SensibilizacionWindow`
- `SeguimientosWindow`
- `SeguimientoEditorWindow`

Nota de arquitectura:

- `seguimientos` sigue siendo un flujo especial. Ya esta formalizado con `supports_drafts=False`, o sea no debo forzarlo dentro del runtime generico de drafts/autosave si una tarea no lo pide expresamente.

## 5. Flujo funcional estandar

Flujo tipico:

1. `HubWindow` abre un formulario.
2. `Section1Window` busca empresa en Supabase y confirma datos base.
3. La ventana del formulario gestiona secciones y autosave.
4. El formulario encola escrituras Supabase usando helpers de `formularios/common.py`.
5. Al finalizar, se construye payload con `completion_payloads.py`.
6. Se escribe en Google Sheets / Excel.
7. Se encola subida a Drive.
8. La UI vuelve al Hub con resultado.

Regla de UI:

- Las tareas costosas deben pasar por `_run_async_ui_task`. No bloquear el hilo principal de Tkinter.

## 6. Persistencia local y rutas

Ubicaciones operativas:

- Logs: `~/Desktop/Inclusion laboral logs/*.log`
- Drafts: `%LOCALAPPDATA%/RECA/drafts.json`
- Formularios completados: `%LOCALAPPDATA%/RECA/completed_forms.json`
- Cola offline Supabase: `%LOCALAPPDATA%/RECA/supabase_write_queue.json`
- Cola Drive: `%LOCALAPPDATA%/RECA/drive_upload_queue.json`
- Credenciales guardadas: `%LOCALAPPDATA%/RECA/login_credentials.json`
- Cache SQLite: `%LOCALAPPDATA%/RECA/cache/offline_cache.db`
- Config Google: `config.json` junto al exe o `%APPDATA%/RECA Inclusion Laboral/`
- Service account: `service-account.json` en rutas equivalentes

## 7. Entorno y dependencias

Dependencias Python declaradas:

- `openpyxl`
- `tkcalendar`
- `pywin32`
- `sounddevice`
- `numpy`
- `google-api-python-client`
- `google-auth`
- `tzdata`
- `pypdf`

Variables relevantes de `.env`:

- `SUPABASE_URL`
- `SUPABASE_KEY`
- `GOOGLE_SERVICE_ACCOUNT_FILE`
- `GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID`
- `GOOGLE_SHEETS_EVALUACION_ACCESIBILIDAD_TEMPLATE_ID`
- `GOOGLE_SHEETS_SEGUIMIENTOS_TEMPLATE_ID`
- `GOOGLE_SHEETS_OFFICIAL_DICTIONARY_SPREADSHEET_ID`
- `SEGUIMIENTOS_SHARED_ROOT`
- `DICTATION_FUNCTION_NAME`
- `DICTATION_LANGUAGE`
- `OPENAI_TEXT_REVIEW_FUNCTION_NAME`
- `INSTALLER_ASSET_NAME`
- `RECA_ENABLE_TEST_FILL`
- `RECA_TEST_FILL_LOGIN`
- `USAGE_EXEMPT_LOGINS`

Notas operativas:

- `GOOGLE_SERVICE_ACCOUNT_FILE` puede ser ruta absoluta o relativa al `.env`.
- La app tambien busca `service-account.json` automaticamente en `%APPDATA%/RECA Inclusion Laboral/` o junto al ejecutable.
- `config.json` y `service-account.json` son recursos runtime criticos para instalaciones reales.

## 8. Integraciones OpenAI / Supabase Edge

Dictado:

- Funcion: `dictate-transcribe`
- Endpoint esperado: `POST {SUPABASE_URL}/functions/v1/dictate-transcribe`
- Auth: `apikey` + `Authorization: Bearer <access_token_usuario>`
- Secrets Supabase:
  - `OPENAI_API_KEY` obligatorio
  - `OPENAI_STT_MODEL` opcional, default `gpt-4o-mini-transcribe`
  - `OPENAI_STT_LANGUAGE` opcional, default `es`
  - limites opcionales de audio y rate limit

Revision ortografica:

- Funcion: `text-review-orthography`
- Endpoint esperado: `POST {SUPABASE_URL}/functions/v1/text-review-orthography`
- Secrets Supabase:
  - `OPENAI_API_KEY` obligatorio
  - `OPENAI_TEXT_REVIEW_MODEL` opcional, default documentado `gpt-4.1-mini`
  - `OPENAI_TEXT_REVIEW_MAX_CHARS` opcional

## 9. Estado reciente del producto

Cambios recientes a tener presentes:

- `2.0.2`:
  - el instalador ya distribuye `service-account.json`
  - el Hub distingue mejor entre "sin conexion" y problemas de configuracion/credenciales
- `2.0.1`:
  - `config.json` ahora debe quedar empaquetado y presente en instalaciones
  - el build valida recursos runtime criticos antes y despues de PyInstaller
- `2.0.0`:
  - refuerzo grande en login, configuracion, cache, actualizacion, feedback visual y resiliencia Sheets/Drive/Seguimientos
- `1.2.7`:
  - mas resiliencia en `seguimientos` y en clasificacion de errores Drive
- `1.2.5`:
  - protecciones contra perdida de datos en formularios de induccion
  - vista `Terminados` para reabrir formularios finalizados recientes
  - `tzdata` agregado para consistencia horaria en Windows

## 10. Hallazgos ya cerrados que no debo reintroducir

Segun la auditoria E2E del 2026-03-26:

- Se corrigio un bloqueo critico en `evaluacion_accesibilidad` entre `section_3 -> section_4`.
- Se unifico el catalogo de `modalidad` para evitar divergencias entre UI y modulo.
- Se corrigio lectura de diccionario en `condiciones_vacante` para no dejar archivos abiertos.
- `seguimientos` quedo con opt-out explicito y testeado del runtime generico de drafts.
- Se ampliaron tests smoke y de regresion; el baseline auditado quedo en `83` tests OK.

Si toco esas zonas, debo validar que no reaparezcan esos problemas.

## 11. Donde mirar cuando falla algo

- Login / auth: `formularios/common.py`, `supabase.log`
- UI congelada: `app.py`, especialmente `_run_async_ui_task`
- Sheets no escribe: `google_sheets_client.py`, mapping del formulario, `excel.log`
- Drive no sube: `drive_upload.py`, `drive.log`, colas persistidas
- Dictado falla: `dictation.py`, `labs.log`, secrets y deploy Edge Function
- Revision ortografica falla: `text_review.py`, `labs.log`, Edge Function
- Finalizacion falla: `completion_payloads.py`, `formularios/finalize_validation.py`, `_finalize_export_flow` en `app.py`
- Datos de empresa incorrectos: cache en `formularios/common.py` y `offline_cache.db`
- Auto-update falla: `updater.py`, `updater.log`
- Problema de encoding al arrancar: `app._run_encoding_health_check`

## 12. Reglas practicas al editar

- No tocar `build/`, `dist/`, `.venv/` ni artefactos generados salvo que la tarea sea de empaquetado o release.
- No cambiar `formularios/common.py`, `google_sheets_client.py`, `drive_upload.py` o `completion_payloads.py` sin revisar impacto transversal.
- En formularios, preferir leer primero el modulo especifico y luego la ventana correspondiente en `app.py`.
- Para agregar un formulario nuevo, partir de `formularios/form_template.py` y seguir `docs/protocolo_nuevo_formulario.md`.
- Si el cambio involucra instalador o release, revisar tambien `README_INSTALL.md`, `build.ps1`, `release.ps1` e `installer.iss`.

## 13. Verificacion minima esperada

Antes de cerrar cambios, elegir lo minimo que corresponda:

- `pytest tests/`
- `python -m unittest discover -s tests -p "test_*.py"`
- smoke local del formulario o flujo tocado
- revisiones de logs relevantes si el cambio es de runtime integrado

Para formularios nuevos o cambios sensibles de UI/export:

- validar cache/resume
- validar export Excel / Sheets
- validar cola o subida a Drive si aplica
- validar retorno al Hub

