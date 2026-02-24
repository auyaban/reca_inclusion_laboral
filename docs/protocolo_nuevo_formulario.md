# Protocolo Estandar para Crear un Formulario

## 1) Estructura de archivos
1. Crear carpeta: `formularios/<nombre_formulario>/`
2. Crear modulo principal: `formularios/<nombre_formulario>/<nombre_formulario>.py`
3. Basarte en `formularios/form_template.py`
4. Confirmar que `register_form()` exista y retorne:
   - `id`
   - `name`

## 2) Registro en Hub (app.py)
1. Importar modulo nuevo.
2. Agregar formulario en `get_forms()`.
3. Crear ventana `*Window` con flujo de secciones.
4. En `_open_form`, agregar caso para abrir la ventana.

## 3) Estandar minimo del modulo
1. Constantes:
   - `FORM_ID`, `FORM_NAME`, `SHEET_NAME`
2. Cache:
   - `FORM_CACHE`, `SECTION_1_CACHE`
   - `save_cache_to_file`, `load_cache_from_file`, `clear_cache_file`
3. Seccion 1 empresa:
   - Buscar por NIT
   - Buscar por nombre
   - Confirmar datos y guardar en cache
4. Confirmadores por seccion:
   - `confirm_section_<n>`
5. Export:
   - `export_to_excel(clear_cache=True)`

## 4) Estandar de salida Excel
1. Carpeta base: `Desktop/Formatos Inclusion Laboral`
2. Subcarpeta por empresa: `<Nombre Empresa>`
3. Archivo:
   - `<Nombre Proceso> - <Nombre Empresa>.xlsx`
4. Reusar template en `templates/`.

## 5) Estandar de integraciones
1. Supabase:
   - Usar helpers de `formularios/common.py`
   - `_supabase_get`
   - `_supabase_upsert` (si aplica)
2. Drive:
   - JSON de respaldo + Excel final en carpetas definidas.

## 6) Estandar de UX
1. Seccion 1 con busqueda NIT + nombre.
2. Reanudar cache si existe.
3. Botones consistentes:
   - Guardar
   - Regresar
   - Continuar/Finalizar
4. Pantalla de carga al finalizar.
5. Retorno a Hub al completar.

## 7) Smoke test obligatorio
1. `python -m py_compile app.py` + modulo nuevo.
2. Flujo completo en GUI con datos minimos.
3. Validar:
   - Cache y resume
   - Export Excel
   - Subida Drive
   - Retorno al Hub
4. Revisar logs:
   - `logs/excel_log.txt`
   - `logs/drive_log.txt`
