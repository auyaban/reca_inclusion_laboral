# Auditoría E2E Main, Formularios y Helpers Compartidos

Fecha: 2026-03-26

## Contexto y cobertura ejecutada
- Rama auditada: `main`
- Alcance auditado: `app.py`, formularios no-Labs (`presentacion_programa`, `evaluacion_accesibilidad`, `condiciones_vacante`, `seleccion_incluyente`, `contratacion_incluyente`, `induccion_organizacional`, `induccion_operativa`, `sensibilizacion`, `seguimientos`) y helpers compartidos de runtime (`formularios/common.py`, `drive_upload.py`, `google_sheets_client.py`, `dictation.py`, `completion_payloads.py`, `logging_utils.py`, `version_info.py`)
- Fuera de alcance: módulos `Labs`, instaladores, release/build y validación contra servicios reales
- Estado del working tree durante la auditoría: `app.py` modificado localmente y `tests/test_evaluacion_accesibilidad_window.py` sin commitear

### Línea base ejecutada
- `python -W default -m unittest discover -s tests -p "test_*.py"`: `83` tests OK
- Sin warnings reproducibles en el baseline posterior al fix
- Import smoke OK para `app`, todos los formularios no-Labs y los helpers compartidos auditados

### Smoke local controlado

| Módulo | Apertura ventana | Render de secciones | Confirmación / navegación | Finalización stub | Resultado |
|---|---|---|---|---|---|
| `presentacion_programa` | OK | OK (`section_1` a `section_5`) | OK en secuencia real | OK | Sin bloqueo reproducible |
| `evaluacion_accesibilidad` | OK | OK | OK en transición real `section_3 -> section_4 -> section_8` | OK con smoke local controlado | Flujo desbloqueado |
| `condiciones_vacante` | OK | OK | OK | OK | Flujo estable en smoke |
| `seleccion_incluyente` | OK | OK | OK | OK | Flujo estable en smoke |
| `contratacion_incluyente` | OK | OK | OK | OK | Flujo estable en smoke |
| `induccion_organizacional` | OK | OK | OK | OK | Flujo estable en smoke |
| `induccion_operativa` | OK | OK | OK | OK | Flujo estable en smoke |
| `sensibilizacion` | OK | OK | OK | OK | Flujo estable en smoke |
| `seguimientos` | OK | OK en ventana inicial y editor stub | OK en save stub del editor | N/A | Flujo especial con opt-out explícito del runtime genérico de drafts |

## Hallazgos priorizados

### `presentacion_programa`
- Sin hallazgos reproducibles de runtime o wiring en esta auditoría.

### `evaluacion_accesibilidad`

#### RECA-AUD-001
- ID: `RECA-AUD-001`
- Módulo: `evaluacion_accesibilidad`
- Flujo afectado: `section_3 -> section_4` en `EvaluacionAccesibilidadWindow`
- Severidad (P0-P3): `P0`
- Tipo: `runtime`
- Evidencia: `app.py:8442-8534`; `app.py:8534` asigna `self._pending_autosave = lambda f=self.section4_fields ...`, pero la clase no crea `self.section4_fields` antes de usarlo. El smoke controlado reprodujo el fallo al ejecutar `_confirm_section_3()` y al renderizar `_show_section_4()`.
- Impacto: el formulario no puede continuar más allá de la sección 3; el usuario queda bloqueado antes de concepto, ajustes razonables, observaciones, cargos compatibles y asistentes.
- Causa probable: refactor incompleto de la sección 4; la sección ahora usa `self.section4_level_var` y `self.section4_desc`, pero dejó un autosave heredado que sigue apuntando a una colección de widgets inexistente.
- Cómo reproducir: abrir `Evaluacion de Accesibilidad`, completar secciones 1, 2.1-2.6 y 3, y luego confirmar la sección 3. La navegación intenta entrar a la sección 4 y lanza `AttributeError: 'EvaluacionAccesibilidadWindow' object has no attribute 'section4_fields'`.
- Recomendación: definir la estructura real de autosave para la sección 4 o reemplazar el lambda por un payload explícito consistente con `_confirm_section_4()`. Añadir un smoke test que cubra la transición `section_3 -> section_4`.
- Estado: `Resuelto` el `2026-03-26`
- Fix aplicado: `app.py` ahora usa `_collect_section4_payload()` tanto en autosave como en `_confirm_section_4()`, eliminando la referencia a `self.section4_fields`.
- Validación posterior: `python -m unittest tests.test_evaluacion_accesibilidad_window -v` y smoke local de navegación `section_3 -> section_4`.

#### RECA-AUD-002
- ID: `RECA-AUD-002`
- Módulo: `evaluacion_accesibilidad`
- Flujo afectado: captura/restauración de `modalidad` en sección 1
- Severidad (P0-P3): `P3`
- Tipo: `wiring`
- Evidencia: `formularios/evaluacion_programa/evaluacion_accesibilidad.py:97-100` define las opciones de módulo como `["Virtual", "Presencial", "Mixto", "No aplica"]`, pero la UI de la ventana en `app.py:9166-9171` inyecta `["Presencial", "Virtual", "Mixta", "No aplica"]`.
- Impacto: el formulario puede guardar una modalidad distinta a la taxonomía del módulo; esto no bloquea el flujo, pero introduce inconsistencia silenciosa entre UI, caché y payload exportado.
- Causa probable: divergencia manual entre catálogo del módulo y catálogo embebido en la clase ventana.
- Cómo reproducir: abrir `Evaluacion de Accesibilidad` y revisar el combo de `Modalidad` de la sección 1; la UI ofrece `Mixta` mientras el módulo declara `Mixto`.
- Recomendación: consolidar el catálogo en una sola fuente de verdad y hacer que la ventana use el listado del módulo.
- Estado: `Resuelto` el `2026-03-26`
- Fix aplicado: `app.py` consume el catálogo de `formularios/evaluacion_programa/evaluacion_accesibilidad.py`, expone alias de compatibilidad `Mixta -> Mixto` en el `Combobox` y el módulo normaliza `modalidad` al confirmar y al cargar caché legacy.
- Validación posterior: `python -m unittest tests.test_evaluacion_accesibilidad_window -v`.

### `condiciones_vacante`

#### RECA-AUD-003
- ID: `RECA-AUD-003`
- Módulo: `condiciones_vacante`
- Flujo afectado: carga local del diccionario de discapacidades
- Severidad (P0-P3): `P3`
- Tipo: `warning`
- Evidencia: `formularios/condiciones_vacante/condiciones_vacante.py:1075-1083` usa `open(...).read()` en dos ramas sin `with`. El baseline `python -W default -m unittest discover ...` reproduce `ResourceWarning: unclosed file` apuntando a `formularios/condiciones_vacante/condiciones_vacante.py:1080`.
- Impacto: no rompe el flujo normal, pero deja descriptores abiertos y mete ruido en baseline/CI; además oculta futuros warnings reales.
- Causa probable: lectura rápida del fallback `Diccionario.txt` sin cierre explícito del handle.
- Cómo reproducir: correr `python -W default -m unittest discover -s tests -p "test_*.py"`; el warning aparece durante los tests de diccionario/voz de `condiciones_vacante`.
- Recomendación: reemplazar ambas lecturas por `with open(...) as handle:` y cubrir el loader con un test que falle si reaparece el warning.
- Estado: `Resuelto` el `2026-03-26`
- Fix aplicado: `formularios/condiciones_vacante/condiciones_vacante.py` usa `with open(...)` en ambas ramas del fallback `utf-8 -> latin-1`.
- Validación posterior: `python -m unittest tests.test_condiciones_vacante_dictionary -v` y `python -W default -m unittest discover -s tests -p "test_*.py"` sin `ResourceWarning`.

### `seleccion_incluyente`
- Sin hallazgos reproducibles de runtime o wiring en esta auditoría.

### `contratacion_incluyente`
- Sin hallazgos reproducibles de runtime o wiring en esta auditoría.

### `induccion_organizacional`
- Sin hallazgos reproducibles de runtime o wiring en esta auditoría.

### `induccion_operativa`
- Sin hallazgos reproducibles de runtime o wiring en esta auditoría.

### `sensibilizacion`
- Sin hallazgos reproducibles de runtime o wiring en esta auditoría.

### `seguimientos`

#### RECA-AUD-004
- ID: `RECA-AUD-004`
- Módulo: `seguimientos`
- Flujo afectado: integración con runtime genérico del hub (drafts, autosave, resolución de módulo)
- Severidad (P0-P3): `P2`
- Tipo: `estructura`
- Evidencia: `get_forms()` sí registra `seguimientos` en `app.py:3720-3732`; `WINDOW_CLASS_FORM_ID_MAP` también lo reconoce en `app.py:164-176`; sin embargo `FORM_MODULE_MAP` en `app.py:152-163` no incluye `seguimientos`. El runtime compartido consulta `FORM_MODULE_MAP` en `app.py:6466-6477` y `app.py:7092-7104`.
- Impacto: `SeguimientosWindow` queda fuera del contrato genérico del hub. No hereda guardado de borrador ni resolución de módulo para features comunes, y cualquier mejora futura que dependa de `FORM_MODULE_MAP` no se aplicará a este flujo.
- Causa probable: `seguimientos` nació como flujo especial y quedó integrado por excepción en `_open_form()`, pero no se completó su registro en la capa genérica.
- Cómo reproducir: abrir `Seguimientos` y revisar el wiring compartido; el formulario abre, pero `_bind_form_runtime()` recibe `module = None` para ese `form_id`, dejando `window._save_draft_command = None`.
- Recomendación: decidir explícitamente si `seguimientos` debe seguir fuera del contrato genérico o incorporarse a `FORM_MODULE_MAP` con una interfaz mínima documentada. Si seguirá siendo especial, aislar esa excepción y documentarla en el código.
- Estado: `Resuelto` el `2026-03-26`
- Fix aplicado: se formalizó el metadata `supports_drafts`; `seguimientos.register_form()` declara `supports_drafts=False` y `app.py` consulta ese contrato en `_get_draft_save_command()`, `_persist_form_draft()`, `_open_draft_entry()`, `_bind_form_runtime()`, `_install_form_autosave_bindings()` y `_schedule_window_draft_autosave()`.
- Validación posterior: `python -m unittest tests.test_seguimientos_runtime -v` y smoke local de apertura de `SeguimientosWindow`.

### Helpers compartidos
- `formularios/common.py`, `drive_upload.py`, `google_sheets_client.py`, `dictation.py`, `completion_payloads.py`, `logging_utils.py` y `version_info.py` cargan correctamente y no mostraron bloqueos de import/runtime en el alcance local de esta auditoría.
- No se validaron integraciones online reales; cualquier riesgo asociado a Supabase, Drive o Sheets queda fuera de este corte.

### Tests y red de regresión

#### RECA-AUD-005
- ID: `RECA-AUD-005`
- Módulo: `tests` / cobertura transversal
- Flujo afectado: regresión de formularios especiales y navegación avanzada
- Severidad (P0-P3): `P2`
- Tipo: `tests`
- Evidencia: el suite actual cubre `common`, plantillas, payloads y algunos helpers, pero no hay smoke UI para `SeguimientosWindow` / `SeguimientoEditorWindow`; en `seguimientos` solo aparece un test de path de plantilla en `tests/test_resource_paths.py:54-63`. Para `evaluacion_accesibilidad`, la prueba existente `tests/test_evaluacion_accesibilidad_window.py:10-16` valida `_maybe_resume_form`, pero no la navegación real `section_3 -> section_4`.
- Impacto: el baseline verde no detectó el bloqueo crítico de `evaluacion_accesibilidad` ni la asimetría estructural de `seguimientos`.
- Causa probable: crecimiento del runtime por ventanas y flujos especiales sin ampliar la red de smoke tests del hub.
- Cómo reproducir: correr el suite actual; obtiene `76` tests OK aun cuando el smoke de UI reproduce el bloqueo de `evaluacion_accesibilidad`.
- Recomendación: añadir smoke tests mínimos de ventana/navegación para `evaluacion_accesibilidad` y `seguimientos` con stubs locales, sin depender de Supabase/Drive reales.
- Estado: `Resuelto` el `2026-03-26`
- Fix aplicado: se extendió `tests/test_evaluacion_accesibilidad_window.py`, `tests/test_condiciones_vacante_dictionary.py` y se agregó `tests/test_seguimientos_runtime.py`.
- Validación posterior: la suite subió de `76` a `83` pruebas y sigue verde con `python -W default -m unittest discover -s tests -p "test_*.py"`.

## Estado posterior al fix
1. `RECA-AUD-001` a `RECA-AUD-005` quedaron resueltos en `main` el `2026-03-26`.
2. La verificación posterior dejó `83` tests OK y smoke cruzado de apertura para todos los formularios no-Labs auditados.
3. `seguimientos` sigue siendo un flujo especial, pero ahora con opt-out explícito y testeado del sistema genérico de borradores/autosave.
