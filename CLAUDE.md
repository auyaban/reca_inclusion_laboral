# CLAUDE.md — Guía de arquitectura para Claude Code

Este archivo se carga automáticamente al inicio de cada conversación de Claude Code.
Úsalo para orientarte antes de leer cualquier archivo fuente.

---

## ¿Qué es este proyecto?

Aplicación de escritorio Windows (Python + Tkinter) para gestionar formularios de
inclusión laboral. Permite diligenciar, guardar y publicar actas de visita a empresas.
Los datos se persisten en **Supabase** (base de datos + auth) y los archivos se suben
a **Google Drive / Sheets**.

Se distribuye como ejecutable compilado con **PyInstaller**.

---

## Mapa de archivos fuente

| Archivo | Responsabilidad |
|---|---|
| `app.py` | Punto de entrada y TODAS las ventanas Tkinter (20 k líneas). Ver sección "Ventanas en app.py" |
| `formularios/common.py` | Utilidades compartidas: Supabase HTTP, rutas de sistema, caché de empresa, queue de escritura offline |
| `formularios/form_template.py` | Plantilla comentada para crear un formulario nuevo. No se importa en producción |
| `formularios/ui_feedback.py` | Sistema de feedback visual en campos (errores inline, banners, placeholders) |
| `formularios/user_messages.py` | Mapa de excepciones → mensajes en español para el usuario final |
| `formularios/finalize_validation.py` | `ValidationIssue` y `format_issues_for_message` — validación pre-envío |
| `completion_payloads.py` | Construcción y serialización de los payloads de finalización de formulario |
| `drive_upload.py` | Worker en hilo de fondo para subir archivos a Google Drive. Cola persistida en JSON |
| `google_sheets_client.py` | Escritura/lectura de Google Sheets con service account. Caché del servicio |
| `google_api_requests.py` | Reintentos con backoff exponencial para Google API. Sin lógica de negocio |
| `dictation.py` | Grabación de audio → transcripción vía Supabase Edge Function (OpenAI STT) |
| `text_review.py` | Revisión ortográfica de campos de texto vía Edge Function (GPT-4.1-nano) |
| `updater.py` | Auto-actualización: compara versión en GitHub Releases, descarga installer |
| `logging_utils.py` | Escribe logs a archivos en Desktop/`Inclusion laboral logs/`. Canales: app, excel, drive, supabase, labs |
| `version_info.py` | Constantes de versión y rutas de datos de la app |

### Formularios (cada uno en `formularios/<nombre>/`)

Cada formulario sigue el mismo patrón: un módulo Python con constantes de mapeo a
columnas de Google Sheets, funciones de construcción de payload, y la ventana
correspondiente en `app.py`.

| Módulo | Ventana en app.py |
|---|---|
| `formularios/presentacion_programa/` | — (sin ventana propia; integrado en HubWindow) |
| `formularios/evaluacion_programa/` | `EvaluacionAccesibilidadWindow` (línea ~10055) |
| `formularios/condiciones_vacante/` | `CondicionesVacanteWindow` (línea ~12028) |
| `formularios/seleccion_incluyente/` | `SeleccionIncluyenteWindow` (línea ~13369) |
| `formularios/contratacion_incluyente/` | `ContratacionIncluyenteWindow` (línea ~14771) |
| `formularios/induccion_organizacional/` | `InduccionOrganizacionalWindow` (línea ~15966) |
| `formularios/induccion_operativa/` | `InduccionOperativaWindow` (línea ~16794) |
| `formularios/sensibilizacion/` | `SensibilizacionWindow` (línea ~17724) |
| `formularios/seguimientos/` | `SeguimientosWindow` (línea ~18156) |

---

## Ventanas en app.py

Las clases de ventana heredan de `tk.Toplevel` + `FormMousewheelMixin`.
Líneas de referencia (aproximadas, buscar con grep):

| Clase | Línea aprox. |
|---|---|
| `HubWindow` (ventana principal / menú) | 6687 |
| `Section1Window` (búsqueda empresa, sección 1 compartida) | 5874 |
| `EvaluacionAccesibilidadWindow` | 10055 |
| `CondicionesVacanteWindow` | 12028 |
| `CondicionesVacanteLabsWindow` | 13315 |
| `SeleccionIncluyenteWindow` | 13369 |
| `SeleccionIncluyenteLabsWindow` | 14761 |
| `ContratacionIncluyenteWindow` | 14771 |
| `InduccionOrganizacionalWindow` | 15966 |
| `InduccionOperativaWindow` | 16794 |
| `SensibilizacionWindow` | 17724 |
| `SeguimientosWindow` | 18156 |
| `SeguimientoEditorWindow` | 19192 |
| `LoadingDialog` | 5010 |
| `LabsSection2VoiceDialog` | 5068 |

Las funciones helper de app.py (~líneas 226–5873) son todas **module-level**, no
métodos de clase. Prefijos de nombre indican su dominio:

- `_autosave_*` / `_guard_*` / `_draft_*` — lógica de guardado automático y borradores
- `_section1_*` — búsqueda y confirmación de empresa (sección 1 compartida)
- `_wizard_*` / `_init_wizard_*` — barra de progreso por secciones
- `_run_async_ui_task` — ejecuta tarea en hilo, muestra LoadingDialog, devuelve al hilo UI
- `_finalize_export_flow` — flujo completo de publicación de acta (Sheets + Drive + PDF)
- `_dpapi_encrypt/decrypt_text` — cifrado de credenciales guardadas (Windows DPAPI)
- `_drive_upload_*` / `_enqueue_pdf_export_job` — cola de subida a Drive

---

## Flujo típico de un formulario

```
HubWindow → selecciona formulario
  → Section1Window (busca empresa en Supabase, confirma sección 1)
    → <FormWindow> (secciones 2..N en wizard)
      → _confirm_section_and_continue() → _supabase_enqueue_upsert() [autosave]
        → _guard_form_finalization() → completion_payloads.build_*()
          → google_sheets_client.write_*() [escribe Sheets]
            → drive_upload._enqueue_drive_upload_job() [sube PDF en background]
              → HubWindow (muestra resultado)
```

---

## Persistencia y datos en disco

| Propósito | Ubicación |
|---|---|
| Logs (5 canales) | `~/Desktop/Inclusion laboral logs/*.log` |
| Borradores de formulario | `%LOCALAPPDATA%/RECA/drafts.json` |
| Formularios completados (índice local) | `%LOCALAPPDATA%/RECA/completed_forms.json` |
| Cola de escritura Supabase (offline) | `%LOCALAPPDATA%/RECA/supabase_write_queue.json` |
| Cola de subida Drive | `%LOCALAPPDATA%/RECA/drive_upload_queue.json` |
| Credenciales guardadas (DPAPI) | `%LOCALAPPDATA%/RECA/login_credentials.json` |
| Caché Supabase (SQLite) | `%LOCALAPPDATA%/RECA/cache/offline_cache.db` |
| Config Google | `config.json` junto al .exe, o `%APPDATA%/RECA Inclusion Laboral/` |
| Service account Google | `service-account.json` (mismos candidatos que config.json) |

---

## Flujos de debugging frecuentes

### El formulario no guarda en Sheets
1. Revisar `~/Desktop/Inclusion laboral logs/excel.log`
2. Buscar en `google_sheets_client.py` la función que falla
3. Verificar que `service-account.json` existe y tiene permisos en la hoja

### El PDF no sube a Drive
1. Revisar `drive.log`
2. Inspeccionar `drive_upload_queue.json` y `drive_failed_queue.json`
3. Función clave: `drive_upload._perform_drive_upload_attempt` (línea ~1286)

### Error de autenticación / login
1. Revisar `supabase.log`
2. Función: `formularios/common._supabase_auth_password_login`
3. Credenciales cifradas: `app._dpapi_encrypt/decrypt_text` (línea ~1450)

### Falla la transcripción de voz
1. Revisar `labs.log`
2. Módulo: `dictation.py` — verifica JWT y URL del Edge Function en `.env`

### Crash al arrancar / encoding
1. `app._run_encoding_health_check()` (línea ~2876) — detecta mojibake en archivos
2. Revisar que todos los `.py` están en UTF-8

### La ventana muestra datos de empresa incorrectos
1. Caché en memoria: `formularios/common._find_cached_company_row`
2. Caché SQLite: `offline_cache.db` tabla `supabase_get_cache`

---

## Contexto mínimo por tarea (lee solo esto)

Antes de explorar el proyecto, identifica tu tarea y lee solo estos archivos:

| Tarea | Archivos a leer |
|---|---|
| Agregar/modificar un campo en un formulario | `formularios/<form>/<form>.py` únicamente |
| Bug al escribir en Google Sheets | `google_sheets_client.py` + `formularios/<form>/<form>.py` |
| Bug al subir a Google Drive | `drive_upload.py` (sección "cola de subida", línea ~912) + `drive.log` |
| Error de autenticación / login | `formularios/common.py` (sección "HTTP Supabase") + `supabase.log` |
| Cambiar un mensaje de error al usuario | `formularios/user_messages.py` (único archivo, 78 líneas) |
| Bug en validación antes de finalizar | `formularios/finalize_validation.py` + `formularios/<form>/<form>.py:validate_before_finalize` |
| Bug en feedback visual de campos | `formularios/ui_feedback.py` (único archivo) |
| Error de credenciales guardadas | `app.py` sección "DPAPI" (buscar `# ── HELPERS: Credenciales`) |
| Problema con borradores (drafts) | `app.py` sección "Borradores" (buscar `# ── HELPERS: Borradores`) |
| Bug en flujo de finalización (Sheets+Drive+PDF) | `app.py:_finalize_export_flow` + `app.py:_start_background_finalization` |
| Problemas con dictado de voz | `dictation.py` + `labs.log` |
| Problemas con revisión ortográfica | `text_review.py` + `labs.log` |
| Bug en auto-actualización | `updater.py` + `updater.log` en Desktop |
| Agregar un formulario nuevo | `formularios/form_template.py` (plantilla) + `app.py:get_forms` |
| Problema de encoding al arrancar | `app.py:_run_encoding_health_check` (buscar con grep) |
| Seguimientos: bug en carpeta/case de Drive | `formularios/seguimientos/seguimientos.py` + `drive.log` |

**Tip:** Para encontrar una sección en `app.py` o `common.py` usa grep sobre el marcador:
```
grep -n "# ── HELPERS: Borradores" app.py
grep -n "# ── HTTP Supabase" formularios/common.py
```

---

## Convenciones de código

- **Prefijo `_`**: función privada de módulo, no forma parte de API pública
- **`_supabase_*`** en `common.py`: todas las llamadas HTTP a Supabase
- **`_supabase_enqueue_*`**: escritura asíncrona con cola offline (preferir sobre `_supabase_upsert` directo)
- **`FORM_CACHE` / `SECTION_1_CACHE`**: dicts globales por módulo de formulario para estado en memoria
- **`form_template.py`**: punto de partida para agregar un formulario nuevo
- Funciones asíncronas de UI siempre pasan por `_run_async_ui_task` — nunca bloquear el hilo principal

---

## Tests

- `tests/` — unitarios con pytest, corren sin Google ni Supabase reales
- `scripts/e2e/` — pruebas end-to-end contra entorno real (requieren `.env` configurado)
- `scripts/smoke_*.py` — smoke tests rápidos por formulario

Correr tests: `pytest tests/`
