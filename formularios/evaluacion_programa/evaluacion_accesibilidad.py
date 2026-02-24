import os
import re
import shutil
import json
import time

from . import seccion_2_4
from . import seccion_2_5_2_6
from . import seccion_3
from . import seccion_4
from . import seccion_5
from . import seccion_6_7
from . import seccion_8
from formularios.common import (
    _get_desktop_dir,
    _normalize_text,
    _sanitize_filename,
    _supabase_get,
)

FORM_NAME = "Evaluacion de Accesibilidad"
SHEET_NAME = "2. EVALUACIÓN DE ACCESIBILIDAD"

SECTIONS = [
    "1. DATOS DE LA EMPRESA",
    "2. ACCESIBILIDAD FÍSICA",
    "2.1 CONDICIONES DE MOVILIDAD Y URBANÍSTICAS",
    "2.2 CONDICIONES DE ACCESIBILIDAD GENERAL",
    "2.3 CONDICIONES DE ACCESIBILIDAD DISCAPACIDAD FÍSICA",
    "2.4 CONDICIONES DE ACCESIBILIDAD DISCAPACIDAD SENSORIAL (VISUAL-AUDITIVA)",
    "2.5 CONDICIONES DE ACCESIBILIDAD DISCAPACIDAD INTELECTUAL - TEA (TRASTORNO ESPECTRO AUTISTA)",
    "2.6 CONDICIONES DE ACCESIBILIDAD DISCAPACIDAD PSICOSOCIAL",
    "3. CONDICIONES ORGANIZACIONALES",
    "4. CONCEPTO DE LA EVALUACIÓN",
    "5. AJUSTES RAZONABLES",
    "6. OBSERVACIONES",
    "7. CARGOS COMPATIBLES",
    "8. ASISTENTES",
]

SECTION_1 = {
    "title": "1. DATOS DE LA EMPRESA",
    "nit_lookup_field": "nit_empresa",
    "fields": [
        {"id": "fecha_visita", "label": "Fecha de la visita", "source": "input"},
        {
            "id": "nombre_empresa",
            "label": "Nombre de la empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "direccion_empresa",
            "label": "Dirección de la empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "correo_1",
            "label": "Correo electrónico",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "contacto_empresa",
            "label": "Contacto de la empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "caja_compensacion",
            "label": "Empresa afiliada a Caja de Compensación",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "asesor",
            "label": "Asesor",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "modalidad",
            "label": "Modalidad",
            "source": "input",
            "options": ["Virtual", "Presencial", "Mixto", "No aplica"],
        },
        {
            "id": "ciudad_empresa",
            "label": "Ciudad/Municipio",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {"id": "nit_empresa", "label": "Número de NIT", "source": "input"},
        {
            "id": "telefono_empresa",
            "label": "Teléfonos",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "cargo",
            "label": "Cargo",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "sede_empresa",
            "label": "Sede Compensar",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "profesional_asignado",
            "label": "Profesional asignado RECA",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
    ],
}

SECTION_1_SUPABASE_MAP = {
    "nombre_empresa": "nombre_empresa",
    "direccion_empresa": "direccion_empresa",
    "correo_1": "correo_1",
    "contacto_empresa": "contacto_empresa",
    "caja_compensacion": "caja_compensacion",
    "asesor": "asesor",
    "ciudad_empresa": "ciudad_empresa",
    "telefono_empresa": "telefono_empresa",
    "cargo": "cargo",
    "sede_empresa": "sede_empresa",
    "profesional_asignado": "profesional_asignado",
}

FORM_CACHE = {}
SECTION_1_CACHE = {}


def _get_cache_dir():
    base = os.getenv("LOCALAPPDATA")
    if not base:
        userprofile = os.getenv("USERPROFILE")
        if userprofile:
            base = os.path.join(userprofile, "AppData", "Local")
    if not base:
        base = os.getcwd()
    cache_dir = os.path.join(base, "RECA", "cache")
    os.makedirs(cache_dir, exist_ok=True)
    return cache_dir


def _get_cache_path():
    return os.path.join(_get_cache_dir(), "evaluacion_accesibilidad.json")


def cache_file_exists():
    return os.path.exists(_get_cache_path())


def save_cache_to_file():
    payload = {
        "form_id": "evaluacion_accesibilidad",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "data": FORM_CACHE,
    }
    with open(_get_cache_path(), "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def load_cache_from_file():
    path = _get_cache_path()
    if not os.path.exists(path):
        return False
    with open(path, "r", encoding="utf-8") as handle:
        payload = json.load(handle) or {}
    data = payload.get("data") or {}
    FORM_CACHE.clear()
    FORM_CACHE.update(data)
    section_1 = data.get("section_1") or {}
    SECTION_1_CACHE.clear()
    SECTION_1_CACHE.update(section_1)
    return True


def clear_cache_file():
    path = _get_cache_path()
    if os.path.exists(path):
        os.remove(path)


def clear_form_cache():
    FORM_CACHE.clear()
    SECTION_1_CACHE.clear()


def _find_template_path():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    templates_dir = os.path.join(base_dir, "templates")
    if not os.path.isdir(templates_dir):
        raise FileNotFoundError("No existe la carpeta templates.")
    for name in os.listdir(templates_dir):
        normalized = _normalize_text(name).replace("_", "")
        if (
            "evaluacion" in normalized
            and "accesibilidad" in normalized
            and normalized.endswith(".xlsx")
        ):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de evaluacion.")


def _ensure_output_path():
    output_path = FORM_CACHE.get("_output_path")
    if output_path and os.path.exists(output_path):
        return output_path
    template_path = _find_template_path()
    desktop = _get_desktop_dir()
    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    safe_company = _sanitize_filename(empresa_nombre)
    if not safe_company:
        safe_company = "Empresa"
    output_dir = os.path.join(desktop, "Formatos Inclusion Laboral", safe_company)
    os.makedirs(output_dir, exist_ok=True)
    process_name = "Evaluacion de Accesibilidad"
    output_name = f"{process_name} - {safe_company}.xlsx"
    output_path = os.path.join(output_dir, output_name)
    if not os.path.exists(output_path):
        shutil.copy2(template_path, output_path)
    FORM_CACHE["_output_path"] = output_path
    return output_path


def get_output_path():
    output_path = FORM_CACHE.get("_output_path")
    if output_path and os.path.exists(output_path):
        return output_path
    return None


def _get_log_dir():
    output_path = FORM_CACHE.get("_output_path")
    if output_path:
        base_dir = os.path.dirname(output_path)
    else:
        base_dir = os.getcwd()
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir


def _log_excel(message):
    try:
        log_dir = _get_log_dir()
        log_path = os.path.join(log_dir, "excel_log.txt")
        reset_log = False
        if os.path.exists(log_path):
            try:
                if os.path.getsize(log_path) >= 5 * 1024 * 1024:
                    reset_log = True
            except OSError:
                reset_log = True
        if reset_log:
            with open(log_path, "w", encoding="utf-8") as log_file:
                log_file.write("")
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] {message}\n")
    except OSError:
        return


def _write_section_with_ws(ws, section_id, payload):
    mapping = EXCEL_MAPPING.get(section_id)
    if not mapping:
        return
    if section_id == "section_8":
        start_row = mapping["start_row"]
        name_col = mapping["name_col"]
        cargo_col = mapping["cargo_col"]
        label_name_col = mapping["label_name_col"]
        label_cargo_col = mapping["label_cargo_col"]
        base_rows = mapping.get("base_rows", 4)
        asistentes = payload or []
        total = len(asistentes)
        template_row = start_row + base_rows - 1
        if total > base_rows:
            insert_at = start_row + base_rows
            extra_rows = total - base_rows
            for _ in range(extra_rows):
                ws.Rows(insert_at).Insert()
                ws.Rows(template_row).Copy()
                ws.Rows(insert_at).PasteSpecial(-4122)
                insert_at += 1
        for idx, entry in enumerate(asistentes):
            row = start_row + idx
            nombre = entry.get("nombre", "")
            cargo = entry.get("cargo", "")
            _log_excel(
                f"WRITE section=section_8 cell={name_col}{row} key=nombre value={nombre!r}"
            )
            _log_excel(
                f"WRITE section=section_8 cell={cargo_col}{row} key=cargo value={cargo!r}"
            )
            ws.Range(f"{name_col}{row}").Value = nombre
            ws.Range(f"{cargo_col}{row}").Value = cargo
            if idx >= base_rows:
                ws.Range(f"{label_name_col}{row}").Value = "Nombre completo:"
                ws.Range(f"{label_cargo_col}{row}").Value = "Cargo:"
    else:
        for key, cell in mapping.items():
            if key in payload:
                value = payload.get(key)
                _log_excel(
                    f"WRITE section={section_id} cell={cell} key={key} value={value!r}"
                )
                ws.Range(cell).Value = value




def export_to_excel(progress_callback=None):
    output_path = _ensure_output_path()
    _log_excel(f"START export_all output={output_path}")
    try:
        import win32com.client as win32
    except ImportError as exc:
        _log_excel("ERROR export_all error=pywin32_not_installed")
        raise RuntimeError("pywin32 no esta instalado. Instala con pip install pywin32.") from exc
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(output_path)
        ws = wb.Worksheets(SHEET_NAME)
        for section_id in EXCEL_MAPPING.keys():
            payload = FORM_CACHE.get(section_id, {})
            _log_excel(f"SECTION export_all section={section_id}")
            if progress_callback:
                progress_callback(section_id)
            _write_section_with_ws(ws, section_id, payload)
        wb.Save()
        _log_excel("SUCCESS export_all")
    except Exception as exc:
        _log_excel(f"ERROR export_all error={exc!r}")
        raise
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
    clear_cache_file()
    clear_form_cache()
    return output_path

EXCEL_MAPPING = {
    "section_1": {
        "fecha_visita": "D7",
        "modalidad": "P7",
        "nit_empresa": "P9",
        "nombre_empresa": "D8",
        "direccion_empresa": "D9",
        "correo_1": "D10",
        "contacto_empresa": "D11",
        "caja_compensacion": "D12",
        "asesor": "D13",
        "ciudad_empresa": "P8",
        "telefono_empresa": "P10",
        "cargo": "P11",
        "sede_empresa": "P12",
        "profesional_asignado": "P13",
    },
    "section_2_1": {
        "transporte_publico_accesible": "M17",
        "transporte_publico_observaciones": "Q17",
        "rutas_pcd_accesible": "M19",
        "rutas_pcd_observaciones": "Q19",
        "parqueaderos_accesible": "M20",
        "parqueaderos_observaciones": "Q20",
        "ubicacion_accesible_accesible": "M21",
        "ubicacion_accesible_observaciones": "Q21",
        "vias_cercanas_accesible": "M22",
        "vias_cercanas_observaciones": "Q22",
        "paso_peatonal_accesible": "M23",
        "paso_peatonal_observaciones": "Q23",
        "rampas_cerca_accesible": "M24",
        "rampas_cerca_observaciones": "Q24",
        "senales_podotactiles": "P25",
        "senales_podotactiles_accesible": "M25",
        "alumbrado_publico": "P26",
        "alumbrado_publico_accesible": "M26",
    },
    "section_2_2": {
        "areas_administrativa_operativa_accesible": "M28",
        "areas_administrativa_operativa": "P28",
        "zonas_comunes_accesible": "M29",
        "zonas_comunes_observaciones": "Q29",
        "enfermeria_accesible_accesible": "M30",
        "enfermeria_accesible": "P30",
        "enfermeria_accesible_observaciones": "Q31",
        "ergonomia_administrativa_accesible": "M32",
        "ergonomia_administrativa": "P32",
        "ergonomia_administrativa_observaciones": "Q33",
        "ergonomia_operativa_accesible": "M34",
        "ergonomia_operativa": "P34",
        "ergonomia_operativa_observaciones": "Q34",
        "mobiliario_zonas_comunes_accesible": "M36",
        "mobiliario_zonas_comunes": "P36",
        "mobiliario_zonas_comunes_observaciones": "Q37",
        "evaluacion_ergonomica_puestos_accesible": "M38",
        "evaluacion_ergonomica_puestos": "P38",
        "ventilacion_area_administrativa_accesible": "M39",
        "ventilacion_area_administrativa": "P39",
        "ventilacion_area_administrativa_secundaria": "P40",
        "ventilacion_area_operativa_accesible": "M41",
        "ventilacion_area_operativa": "P41",
        "ventilacion_area_operativa_secundaria": "P42",
        "ventilacion_areas_comunes_accesible": "M43",
        "ventilacion_areas_comunes": "P43",
        "ventilacion_areas_comunes_secundaria": "P44",
        "iluminacion_area_administrativa_accesible": "M45",
        "iluminacion_area_administrativa": "P45",
        "iluminacion_area_administrativa_secundaria": "P46",
        "iluminacion_area_operativa_accesible": "M47",
        "iluminacion_area_operativa": "P47",
        "iluminacion_area_operativa_secundaria": "P48",
        "iluminacion_areas_comunes_accesible": "M49",
        "iluminacion_areas_comunes": "P49",
        "iluminacion_areas_comunes_secundaria": "P50",
        "ruido_area_administrativa_accesible": "M51",
        "ruido_area_administrativa": "P51",
        "ruido_area_administrativa_secundaria": "P52",
        "ruido_area_administrativa_terciaria": "P53",
        "ruido_area_operativa_accesible": "M54",
        "ruido_area_operativa": "P54",
        "ruido_area_operativa_secundaria": "P55",
        "ruido_area_operativa_terciaria": "P56",
        "ruido_areas_comunes_accesible": "M57",
        "ruido_areas_comunes": "P57",
        "ruido_areas_comunes_secundaria": "P58",
        "ruido_areas_comunes_terciaria": "P59",
        "flexibilidad_hibrido_remoto_accesible": "M60",
        "flexibilidad_hibrido_remoto": "P60",
        "flexibilidad_horarios_calamidades_accesible": "M61",
        "flexibilidad_horarios_calamidades": "P61",
        "sala_lactancia_accesible": "M62",
        "sala_lactancia": "P62",
        "protocolo_sala_lactancia_accesible": "M63",
        "protocolo_sala_lactancia": "P63",
        "linea_purpura_accesible": "M64",
        "linea_purpura": "Q64",
        "salas_amigas_accesible": "M65",
        "salas_amigas": "P65",
        "protocolo_hostigamiento_acoso_sexual_accesible": "M66",
        "protocolo_hostigamiento_acoso_sexual": "Q66",
        "protocolo_acoso_laboral_accesible": "M67",
        "protocolo_acoso_laboral": "Q67",
        "practicas_equidad_genero_accesible": "M68",
        "practicas_equidad_genero": "Q68",
        "canales_comunicacion_lenguaje_inclusivo_accesible": "M69",
        "canales_comunicacion_lenguaje_inclusivo": "Q69",
    },
    "section_2_3": {
        "entrada_salida_accesible": "M71",
        "entrada_salida_observaciones": "Q71",
        "rampas_interior_usr_accesible": "M72",
        "rampas_interior_usr_observaciones": "Q72",
        "ascensor_interior_accesible": "M73",
        "ascensor_interior_observaciones": "Q73",
        "zonas_oficinas_accesibles_accesible": "M74",
        "zonas_oficinas_accesibles_observaciones": "Q74",
        "cafeteria_accesible_accesible": "M75",
        "cafeteria_accesible_observaciones": "Q75",
        "zonas_descanso_accesibles_accesible": "M76",
        "zonas_descanso_accesibles_observaciones": "Q76",
        "pasillos_amplios_accesible": "M77",
        "pasillos_amplios_observaciones": "Q77",
        "escaleras_doble_funcion_accesible": "M78",
        "escaleras_doble_funcion": "P78",
        "escaleras_doble_funcion_secundaria": "P79",
        "escaleras_doble_funcion_terciaria": "P80",
        "escaleras_doble_funcion_cuaternaria": "P81",
        "escaleras_interior_accesible": "M82",
        "escaleras_interior": "P82",
        "escaleras_interior_secundaria": "P83",
        "escaleras_interior_terciaria": "P84",
        "escaleras_interior_cuaternaria": "P85",
        "escaleras_emergencia_accesible": "M86",
        "escaleras_emergencia": "P86",
        "escaleras_emergencia_secundaria": "P87",
        "escaleras_emergencia_terciaria": "P88",
        "bano_discapacidad_fisica_accesible": "M89",
        "bano_discapacidad_fisica": "P89",
        "bano_discapacidad_fisica_secundaria": "P90",
        "bano_discapacidad_fisica_terciaria": "P91",
        "bano_discapacidad_fisica_cuaternaria": "P92",
        "bano_discapacidad_fisica_quinary": "P93",
        "silla_evacuacion_usr_accesible": "M94",
        "silla_evacuacion_usr": "P94",
        "silla_evacuacion_oruga_accesible": "M95",
        "silla_evacuacion_oruga": "P95",
        "ergonomia_superficies_irregulares_accesible": "M96",
        "ergonomia_superficies_irregulares": "P96",
        "senalizacion_ntc_accesible": "M97",
        "senalizacion_ntc": "P97",
        "senalizacion_ntc_secundaria": "P98",
        "mapa_evacuacion_ntc_accesible": "M99",
        "mapa_evacuacion_ntc": "P99",
        "mapa_evacuacion_ntc_secundaria": "P100",
        "mapa_evacuacion_ntc_terciaria": "P101",
        "ajustes_razonables_individualizados_accesible": "M102",
        "ajustes_razonables_individualizados": "P102",
        "ajustes_razonables_detalle": "Q103",
    },
    "section_2_4": {
        "senalizacion_orientacion_accesible": "M105",
        "senalizacion_orientacion": "P105",
        "senalizacion_emergencia_accesible": "M106",
        "senalizacion_emergencia": "P106",
        "distribucion_zonas_comunes_accesible": "M107",
        "distribucion_zonas_comunes_observaciones": "Q107",
        "senalizacion_mapa_evacuacion_accesible": "M108",
        "senalizacion_mapa_evacuacion_observaciones": "Q108",
        "ascensor_apoyo_visual_sonoro_accesible": "M109",
        "ascensor_apoyo_visual_sonoro_observaciones": "Q109",
        "apoyo_seguridad_ubicacion_accesible": "M110",
        "apoyo_seguridad_ubicacion_observaciones": "Q110",
        "senalizacion_ntc_accesible": "M111",
        "senalizacion_ntc": "P111",
        "senalizacion_ntc_secundaria": "P112",
        "mapa_evacuacion_ntc_accesible": "M113",
        "mapa_evacuacion_ntc": "P113",
        "mapa_evacuacion_ntc_secundaria": "P114",
        "mapa_evacuacion_ntc_terciaria": "P115",
        "informacion_accesible_ingreso_accesible": "M116",
        "informacion_accesible_ingreso": "P116",
        "informacion_accesible_ingreso_secundaria": "P117",
        "informacion_accesible_ingreso_terciaria": "P118",
        "informacion_accesible_ingreso_cuaternaria": "P119",
        "medios_tecnologicos_ingreso_accesible": "M120",
        "medios_tecnologicos_ingreso_observaciones": "Q120",
        "material_seleccion_accesible_accesible": "M121",
        "material_seleccion_accesible": "P121",
        "material_seleccion_accesible_secundaria": "P122",
        "material_seleccion_accesible_terciaria": "P123",
        "material_contratacion_accesible_accesible": "M124",
        "material_contratacion_accesible": "P124",
        "material_contratacion_accesible_secundaria": "P126",
        "material_induccion_accesible_accesible": "M127",
        "material_induccion_accesible": "P127",
        "material_induccion_accesible_secundaria": "P129",
        "material_induccion_accesible_terciaria": "P130",
        "material_evaluacion_desempeno_accesible": "M131",
        "material_evaluacion_desempeno": "P131",
        "material_evaluacion_desempeno_secundaria": "P132",
        "material_evaluacion_desempeno_terciaria": "P133",
        "plataformas_autogestion_accesible": "M134",
        "plataformas_autogestion": "P134",
        "plataformas_autogestion_secundaria": "P135",
        "plataformas_autogestion_terciaria": "P137",
        "plataformas_autogestion_cuaternaria": "P138",
        "alarma_emergencia_accesible": "M139",
        "alarma_emergencia": "P139",
        "ajustes_razonables_individualizados_accesible": "M140",
        "ajustes_razonables_individualizados": "P140",
        "ajustes_razonables_individualizados_detalle": "Q141",
    },
    "section_2_5": {
        "material_seleccion_cognitiva_accesible": "M143",
        "material_seleccion_cognitiva": "P143",
        "material_contratacion_cognitiva_accesible": "M144",
        "material_contratacion_cognitiva": "P144",
        "material_induccion_cognitiva_accesible": "M145",
        "material_induccion_cognitiva": "P145",
        "material_evaluacion_cognitiva_accesible": "M146",
        "material_evaluacion_cognitiva": "P146",
        "ascensor_facil_ubicacion_accesible": "M147",
        "ascensor_facil_ubicacion_observaciones": "Q147",
        "distribucion_zonas_comunes_percepcion_accesible": "M148",
        "distribucion_zonas_comunes_percepcion_observaciones": "Q148",
        "plataformas_autogestion_intelectual_accesible": "M149",
        "plataformas_autogestion_intelectual": "P149",
        "plataformas_autogestion_intelectual_secundaria": "P150",
        "ajustes_razonables_intelectual_accesible": "M151",
        "ajustes_razonables_intelectual": "P151",
        "ajustes_razonables_intelectual_detalle": "Q152",
    },
    "section_2_6": {
        "ajustes_razonables_psicosocial_accesible": "M154",
        "ajustes_razonables_psicosocial": "P154",
        "ajustes_razonables_psicosocial_detalle": "Q155",
    },
    "section_3": {
        "experiencia_vinculacion_pcd_accesible": "M157",
        "experiencia_vinculacion_pcd_observaciones": "Q157",
        "personal_tercerizado_capacitado_accesible": "M158",
        "personal_tercerizado_capacitado": "P158",
        "personal_directo_capacitado_accesible": "M159",
        "personal_directo_capacitado": "P159",
        "apoyo_arl_seguridad_accesible": "M160",
        "apoyo_arl_seguridad": "P160",
        "capacitacion_emergencias_accesible": "M161",
        "capacitacion_emergencias_observaciones": "Q161",
        "politica_diversidad_inclusion_accesible": "M162",
        "politica_diversidad_inclusion": "P162",
        "rrhh_normatividad_accesible": "M163",
        "rrhh_normatividad": "P163",
        "rrhh_normatividad_secundaria": "P164",
        "rrhh_normatividad_terciaria": "P165",
        "rrhh_normatividad_cuaternaria": "P166",
        "ajustes_razonables_empresa_accesible": "M167",
        "ajustes_razonables_empresa": "P167",
        "ajustes_razonables_empresa_secundaria": "P168",
        "ajustes_razonables_empresa_terciaria": "P169",
        "ajustes_razonables_empresa_cuaternaria": "P170",
        "protocolo_emergencias_pcd_accesible": "M171",
        "protocolo_emergencias_pcd": "P171",
        "protocolo_emergencias_pcd_secundaria": "P172",
        "protocolo_emergencias_pcd_terciaria": "P173",
        "apoyo_bomberos_discapacidad_accesible": "M174",
        "apoyo_bomberos_discapacidad": "P174",
        "apoyo_bomberos_discapacidad_detalle": "Q175",
        "disponibilidad_tiempo_inclusion_accesible": "M176",
        "disponibilidad_tiempo_inclusion": "P176",
        "practicas_equidad_genero_accesible": "M177",
        "practicas_equidad_genero": "P177",
    },
    "section_4": {
        "nivel_accesibilidad": "M180",
        "descripcion": "Q180",
    },
    "section_5": {
        "discapacidad_fisica_ajustes": "K186",
        "discapacidad_fisica": "G186",
        "discapacidad_fisica_nota": "A187",
        "discapacidad_fisica_usr_ajustes": "K188",
        "discapacidad_fisica_usr": "G188",
        "discapacidad_fisica_usr_nota": "A189",
        "discapacidad_auditiva_ajustes": "K190",
        "discapacidad_auditiva": "G190",
        "discapacidad_auditiva_nota": "A191",
        "discapacidad_visual_ajustes": "K192",
        "discapacidad_visual": "G192",
        "discapacidad_visual_nota": "A193",
        "discapacidad_intelectual_ajustes": "K194",
        "discapacidad_intelectual": "G194",
        "discapacidad_intelectual_nota": "A195",
        "trastorno_espectro_autista_ajustes": "K196",
        "trastorno_espectro_autista": "G196",
        "trastorno_espectro_autista_nota": "A197",
        "discapacidad_psicosocial_ajustes": "K198",
        "discapacidad_psicosocial": "G198",
        "discapacidad_psicosocial_nota": "A199",
        "discapacidad_visual_baja_vision_ajustes": "K200",
        "discapacidad_visual_baja_vision": "G200",
        "discapacidad_visual_baja_vision_nota": "A201",
        "discapacidad_auditiva_reducida_ajustes": "K202",
        "discapacidad_auditiva_reducida": "G202",
        "discapacidad_auditiva_reducida_nota": "A203",
    },
    "section_6": {
        "observaciones_generales": "A205",
    },
    "section_7": {
        "cargos_compatibles": "A208",
    },
    "section_8": {
        "start_row": 212,
        "name_col": "C",
        "cargo_col": "O",
        "label_name_col": "A",
        "label_cargo_col": "N",
        "base_rows": 4,
    },
}

SECTION_2_1 = {
    "title": "2.1 CONDICIONES DE MOVILIDAD Y URBANÍSTICAS",
    "questions": [
        {
            "id": "transporte_publico",
            "label": "¿Existe transporte público para ingresar y salir de la empresa?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "rutas_pcd",
            "label": "¿La empresa cuenta con rutas para sus vinculados PcD?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "parqueaderos",
            "label": "¿La organización cuenta con parqueaderos? (Vehículos, moto, bicicletas)",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "ubicacion_accesible",
            "label": "¿La ubicación de la empresa es accesible?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "vias_cercanas",
            "label": "¿Las vías cercanas a la empresa son accesibles?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "paso_peatonal",
            "label": (
                "¿Existe un paso peatonal accesible? (Puente peatonal, cebra, "
                "semáforo, sendero peatonal)"
            ),
            "type": "accesible_con_observaciones",
        },
        {
            "id": "rampas_cerca",
            "label": "¿Existen rampas cerca a la empresa?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "senales_podotactiles",
            "label": "¿Existen señales podotáctiles?",
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Presencia de señales podotáctiles continuas y en buen estado.",
                "Presencia de señales podotáctiles discontinuas y en buen estado.",
                "No aplica.",
                "Presencia de señales podotáctiles continuas y en mal estado.",
                "Presencia de señales podotáctiles discontinuas y en mal estado.",
            ],
        },
        {
            "id": "alumbrado_publico",
            "label": "¿Se cuenta con alumbrado público cerca a la empresa?",
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Alumbrado público en buen estado.",
                "Alumbrado público en mal estado.",
                "No aplica",
            ],
        },
    ],
    "accesible_options": ["Sí", "No", "Parcial"],
}

SECTION_2_2 = {
    "title": "2.2 CONDICIONES DE ACCESIBILIDAD GENERAL",
    "questions": [
        {
            "id": "areas_administrativa_operativa",
            "label": "¿La empresa cuenta con área administrativa y operativa?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Cuenta con ambas áreas.",
                "Se cuenta con área operativa.",
                "Se cuenta con área administrativa.",
                "No aplica.",
            ],
        },
        {
            "id": "zonas_comunes",
            "label": "¿La empresa cuenta con zonas comunes? (Describa cuáles)",
            "type": "accesible_con_observaciones",
            "has_accesible": True,
            "has_observaciones": True,
        },
        {
            "id": "enfermeria_accesible",
            "label": "¿La empresa cuenta con enfermería y es accesible para la PcD?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": True,
            "options": [
                "Accesible para toda la población.",
                "No accesible para USR.",
                "No accesible para personas con discapacidad física con apoyo diferente a silla de ruedas.",
                "No accesible para personas con discapacidad visual.",
                "No accesible para personas con discapacidad auditiva.",
                "No accesible para personas con discapacidad cognitiva.",
                "No accesible para personas con TEA.",
                "No accesible para personas con discapacidad psicosocial.",
                "No aplica.",
            ],
        },
        {
            "id": "ergonomia_administrativa",
            "label": (
                "En el área administrativa, el diseño de los puestos de trabajo y mobiliario "
                "es ergonómico y está en buen estado?"
            ),
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": True,
            "options": [
                "Accesible para toda la población.",
                "No accesible para USR.",
                "No accesible para personas con discapacidad física con apoyo diferente a silla de ruedas.",
                "No accesible para personas con discapacidad visual.",
                "No accesible para personas con discapacidad auditiva.",
                "No accesible para personas con discapacidad cognitiva.",
                "No accesible para personas con TEA.",
                "No accesible para personas con discapacidad psicosocial.",
                "No aplica.",
            ],
        },
        {
            "id": "ergonomia_operativa",
            "label": (
                "En el área operativa, el diseño de los puestos de trabajo y mobiliario son "
                "ergonómicos y están en buen estado?"
            ),
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": True,
            "options": [
                "Accesible para toda la población.",
                "No accesible para USR.",
                "No accesible para personas con discapacidad física con apoyo diferente a silla de ruedas.",
                "No accesible para personas con discapacidad visual.",
                "No accesible para personas con discapacidad auditiva.",
                "No accesible para personas con discapacidad cognitiva.",
                "No accesible para personas con TEA.",
                "No accesible para personas con discapacidad psicosocial.",
                "No aplica.",
            ],
        },
        {
            "id": "mobiliario_zonas_comunes",
            "label": "¿En las zonas comunes, el mobiliario se encuentra en buen estado?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": True,
            "options": [
                "Accesible para toda la población.",
                "No accesible para USR.",
                "No accesible para personas con discapacidad física con apoyo diferente a silla de ruedas.",
                "No accesible para personas con discapacidad visual.",
                "No accesible para personas con discapacidad auditiva.",
                "No accesible para personas con discapacidad cognitiva.",
                "No accesible para personas con TEA.",
                "No accesible para personas con discapacidad psicosocial.",
                "No aplica.",
            ],
        },
        {
            "id": "evaluacion_ergonomica_puestos",
            "label": "¿La organización ha realizado evaluación ergonómica de los puestos de trabajo?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Con apoyo de la ARL.",
                "Sin apoyo de la ARL.",
                "No aplica.",
            ],
        },
        {
            "id": "ventilacion_area_administrativa",
            "label": "¿Con qué tipo de ventilación se cuenta en el área administrativa?",
            "type": "lista_doble",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Ventilación natural.",
                "Ventilación artificial.",
                "Ventilación natural y artificial.",
                "No cuenta con ventilación.",
                "No aplica.",
            ],
            "options_secondary": [
                "La organización ha realizado mediciones higiénicas relacionadas con ventilación.",
                "La organización NO ha realizado mediciones higiénicas relacionadas con ventilación.",
                "No aplica.",
            ],
        },
        {
            "id": "ventilacion_area_operativa",
            "label": "¿Con qué tipo de ventilación se cuenta en el área operativa?",
            "type": "lista_doble",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Ventilación natural.",
                "Ventilación artificial.",
                "Ventilación natural y artificial.",
                "No cuenta con ventilación.",
                "No aplica.",
            ],
            "options_secondary": [
                "La organización ha realizado mediciones higiénicas relacionadas con ventilación.",
                "La organización NO ha realizado mediciones higiénicas relacionadas con ventilación.",
                "No aplica.",
            ],
        },
        {
            "id": "ventilacion_areas_comunes",
            "label": "¿Con qué tipo de ventilación se cuenta en las áreas comunes?",
            "type": "lista_doble",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Ventilación natural.",
                "Ventilación artificial.",
                "Ventilación natural y artificial.",
                "No cuenta con ventilación.",
                "No aplica.",
            ],
            "options_secondary": [
                "La organización ha realizado mediciones higiénicas relacionadas con ventilación.",
                "La organización NO ha realizado mediciones higiénicas relacionadas con ventilación.",
                "No aplica.",
            ],
        },
        {
            "id": "iluminacion_area_administrativa",
            "label": "¿Con qué tipo de iluminación se cuenta en el área administrativa?",
            "type": "lista_doble",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Iluminación natural y artificial.",
                "Iluminación artificial.",
                "Iluminación natural.",
                "No aplica.",
            ],
            "options_secondary": [
                "La organización ha realizado mediciones higiénicas en iluminación.",
                "La organización no ha realizado mediciones higiénicas en iluminación.",
                "No aplica.",
            ],
        },
        {
            "id": "iluminacion_area_operativa",
            "label": "¿Con qué tipo de iluminación se cuenta en el área operativa?",
            "type": "lista_doble",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Iluminación natural y artificial.",
                "Iluminación artificial.",
                "Iluminación natural.",
                "No aplica.",
            ],
            "options_secondary": [
                "La organización ha realizado mediciones higiénicas en iluminación.",
                "La organización no ha realizado mediciones higiénicas en iluminación.",
                "No aplica.",
            ],
        },
        {
            "id": "iluminacion_areas_comunes",
            "label": "¿Con qué tipo de iluminación se cuenta en las áreas comunes?",
            "type": "lista_doble",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Iluminación natural y artificial.",
                "Iluminación artificial.",
                "Iluminación natural.",
                "No aplica.",
            ],
            "options_secondary": [
                "La organización ha realizado mediciones higiénicas en iluminación.",
                "La organización no ha realizado mediciones higiénicas en iluminación.",
                "No aplica.",
            ],
        },
        {
            "id": "ruido_area_administrativa",
            "label": "¿El nivel de ruido en el área administrativa es adecuado?",
            "type": "lista_triple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Percepción de ruido bajo.",
                "Percepción de ruido medio.",
                "Percepción de ruido alto.",
                "No aplica.",
            ],
            "options_secondary": [
                "Se requiere el uso de elemento de protección auditiva.",
                "No se requiere el uso de elementos de protección auditiva.",
                "No aplica.",
            ],
            "options_tertiary": [
                "La organización ha realizado mediciones higiénicas de ruido.",
                "La organización No ha realizado mediciones higiénicas de ruido.",
                "No aplica.",
            ],
        },
        {
            "id": "ruido_area_operativa",
            "label": "¿El nivel de ruido en el área operativa es adecuado?",
            "type": "lista_triple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Percepción de ruido bajo.",
                "Percepción de ruido medio.",
                "Percepción de ruido alto.",
                "No aplica.",
            ],
            "options_secondary": [
                "Se requiere el uso de elemento de protección auditiva.",
                "No se requiere el uso de elementos de protección auditiva.",
                "No aplica.",
            ],
            "options_tertiary": [
                "La organización ha realizado mediciones higiénicas de ruido.",
                "La organización No ha realizado mediciones higiénicas de ruido.",
                "No aplica.",
            ],
        },
        {
            "id": "ruido_areas_comunes",
            "label": "¿El nivel de ruido de las áreas comunes es adecuado?",
            "type": "lista_triple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Percepción de ruido bajo.",
                "Percepción de ruido medio.",
                "Percepción de ruido alto.",
                "No aplica.",
            ],
            "options_secondary": [
                "Se requiere el uso de elemento de protección auditiva.",
                "No se requiere el uso de elementos de protección auditiva.",
                "No aplica.",
            ],
            "options_tertiary": [
                "La organización ha realizado mediciones higiénicas de ruido.",
                "La organización No ha realizado mediciones higiénicas de ruido.",
                "No aplica.",
            ],
        },
        {
            "id": "flexibilidad_hibrido_remoto",
            "label": "¿La empresa cuenta con flexibilidad de trabajo híbrido o remoto?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "No cuenta con flexibilidad de trabajo.",
                "Sí cuenta con flexibilidad de trabajo.",
                "No aplica.",
            ],
        },
        {
            "id": "flexibilidad_horarios_calamidades",
            "label": (
                "¿La empresa cuenta con flexibilidad de horarios ante calamidades domésticas, "
                "teniendo en cuenta las mujeres cuidadoras?"
            ),
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "No cuenta con flexibilidad.",
                "Sí cuenta con flexibilidad.",
                "No aplica.",
            ],
        },
        {
            "id": "sala_lactancia",
            "label": "¿La empresa cuenta con sala de lactancia?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": ["Sí", "No", "En construcción"],
        },
        {
            "id": "protocolo_sala_lactancia",
            "label": "¿La empresa cuenta con protocolo de manejo en sala de lactancia según resolución 2423 del 2018?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": ["Sí", "No", "En construcción"],
        },
        {
            "id": "linea_purpura",
            "label": "¿La empresa conoce y maneja la línea Púrpura?",
            "type": "texto",
            "has_accesible": True,
            "has_observaciones": False,
        },
        {
            "id": "salas_amigas",
            "label": "¿La empresa conoce y maneja salas amigas?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": ["Sí", "No", "En construcción"],
        },
        {
            "id": "protocolo_hostigamiento_acoso_sexual",
            "label": "¿La empresa cuenta con un protocolo contra el hostigamiento y acoso sexual?",
            "type": "texto",
            "has_accesible": True,
            "has_observaciones": False,
        },
        {
            "id": "protocolo_acoso_laboral",
            "label": "¿La empresa cuenta con un protocolo contra el acoso laboral?",
            "type": "texto",
            "has_accesible": True,
            "has_observaciones": False,
        },
        {
            "id": "practicas_equidad_genero",
            "label": "¿La empresa cuenta con prácticas inclusivas orientadas a la equidad de género, cuáles?",
            "type": "texto",
            "has_accesible": True,
            "has_observaciones": False,
        },
        {
            "id": "canales_comunicacion_lenguaje_inclusivo",
            "label": (
                "¿Cuenta con canales de comunicación interno que promuevan el uso de lenguaje inclusivo?"
            ),
            "type": "texto",
            "has_accesible": True,
            "has_observaciones": False,
        },
    ],
    "accesible_options": ["Sí", "No", "Parcial"],
}

SECTION_2_3 = {
    "title": "2.3 CONDICIONES DE ACCESIBILIDAD DISCAPACIDAD FÍSICA",
    "questions": [
        {
            "id": "entrada_salida",
            "label": "¿La entrada y salida de la empresa es accesible?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "rampas_interior_usr",
            "label": "¿Hay rampas al interior de la empresa para el acceso a USR?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "ascensor_interior",
            "label": "¿Hay ascensor al interior de las instalaciones?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "zonas_oficinas_accesibles",
            "label": "¿Las zonas de oficinas son accesibles?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "cafeteria_accesible",
            "label": "¿La cafetería es accesible?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "zonas_descanso_accesibles",
            "label": "¿Las zonas de descanso son accesibles?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "pasillos_amplios",
            "label": "¿Los pasillos son amplios y permiten el desplazamiento independiente?",
            "type": "accesible_con_observaciones",
        },
        {
            "id": "escaleras_doble_funcion",
            "label": (
                "¿Las escaleras cumplen una doble función, tanto para emergencias como para "
                "el tránsito interno de la empresa?"
            ),
            "type": "lista_multiple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Cuenta con pasamanos (Altura de 85 a 100 cm).",
                "No cuenta con pasamanos (Altura de 85 a 100 cm).",
                "No aplica.",
            ],
            "options_secondary": [
                "Cuenta con pasamanos a un solo costado de la escalera.",
                "Cuenta con pasamanos en ambos costados.",
                "No aplica.",
            ],
            "options_tertiary": [
                "Cuenta con bandas antideslizantes.",
                "No cuenta con bandas antideslizantes.",
                "No aplica.",
            ],
            "options_quaternary": [
                "Cuenta con las características básicas de un diseño seguro (Ancho mínimo 90 cm).",
                "No cuenta con las características básicas de un diseño seguro (Ancho mínimo 90 cm).",
                "No aplica.",
            ],
        },
        {
            "id": "escaleras_interior",
            "label": "¿Hay escaleras al interior de las instalaciones?",
            "type": "lista_multiple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Cuenta con pasamanos (Altura de 85 a 100 cm).",
                "No cuenta con pasamanos (Altura de 85 a 100 cm).",
                "No aplica.",
            ],
            "options_secondary": [
                "Cuenta con pasamanos a un solo costado de la escalera.",
                "Cuenta con pasamanos en ambos costados.",
                "No aplica.",
            ],
            "options_tertiary": [
                "Cuenta con bandas antideslizantes.",
                "No cuenta con bandas antideslizantes.",
                "No aplica.",
            ],
            "options_quaternary": [
                "Cuenta con las características básicas de un diseño seguro (Ancho mínimo 90 cm).",
                "No cuenta con las características básicas de un diseño seguro (Ancho mínimo 90 cm).",
                "No aplica.",
            ],
        },
        {
            "id": "escaleras_emergencia",
            "label": "¿Cuenta con escaleras de emergencia?",
            "type": "lista_triple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Cuenta con pasamanos (Altura de 85 a 100 cm).",
                "No cuenta con pasamanos (Altura de 85 a 100 cm).",
                "No aplica.",
            ],
            "options_secondary": [
                "Cuenta con bandas antideslizantes.",
                "No cuenta con bandas antideslizantes.",
                "No aplica.",
            ],
            "options_tertiary": [
                "Cuenta con las características básicas de un diseño seguro (Ancho mínimo 90 cm).",
                "No cuenta con las características básicas de un diseño seguro (Ancho mínimo 90 cm).",
                "No aplica.",
            ],
        },
        {
            "id": "bano_discapacidad_fisica",
            "label": "¿Existe un baño para discapacidad física?",
            "type": "lista_multiple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Cuenta con barras de agarre en ambos lados de la unidad sanitaria.",
                "No cuenta con barras de agarre en ambos lados de la unidad sanitaria.",
                "No aplica.",
            ],
            "options_secondary": [
                "Cuenta con espacio mínimo de lado o al frente de 120 cm.",
                "No cuenta con espacio mínimo de lado o al frente de 120 cm.",
                "No aplica.",
            ],
            "options_tertiary": [
                "Cuenta con lavamanos de altura de 75 cm.",
                "No cuenta con lavamanos de altura de 75 cm.",
                "No aplica.",
            ],
            "options_quaternary": [
                "Cuenta con timbre de emergencia situado al lado del sanitario.",
                "No cuenta con timbre de emergencia situado al lado del sanitario.",
                "No aplica.",
            ],
            "options_quinary": [
                "Los accesorios interfieren con las barras de apoyo.",
                "Los accesorios NO interfieren con las barras de apoyo.",
                "No aplica.",
            ],
        },
        {
            "id": "silla_evacuacion_usr",
            "label": "¿La empresa cuenta con silla de evacuación para personas USR?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Silla de evacuación en buen estado.",
                "Silla de evacuación en mal estado.",
                "No aplica.",
            ],
        },
        {
            "id": "silla_evacuacion_oruga",
            "label": "¿La empresa cuenta con silla de evacuación tipo oruga?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Silla tipo oruga en buen estado.",
                "Silla tipo oruga en mal estado.",
                "No aplica.",
            ],
        },
        {
            "id": "ergonomia_superficies_irregulares",
            "label": (
                "¿La organización ha realizado evaluación ergonómica de superficies sin "
                "irregularidades? (Sin desniveles bruscos y con suelos que sean antideslizantes)"
            ),
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Con apoyo de la ARL.",
                "Sin apoyo de la ARL.",
                "No aplica.",
            ],
        },
        {
            "id": "senalizacion_ntc",
            "label": "¿La señalización cumple con los criterios de la NTC?",
            "type": "lista_doble",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "La señalización cumple con los criterios de altura de acuerdo a la NTC (Altura de 120 cm a 160 cm).",
                "La señalización No cumple con los criterios de altura de acuerdo a la NTC (Altura de 120 cm a 160 cm).",
                "No aplica.",
                "No se cuenta con señalización de orientación.",
            ],
            "options_secondary": [
                "La señalización cumple con los criterios de material de acuerdo a la NTC.",
                "La señalización No cumple con los criterios de material de acuerdo a la NTC.",
                "No aplica.",
            ],
        },
        {
            "id": "mapa_evacuacion_ntc",
            "label": "¿El mapa de evacuación cumple con criterios de accesibilidad de la NTC?",
            "type": "lista_triple",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Cuenta con señalización en el área operativa.",
                "Cuenta con señalización en el área administrativa.",
                "Cuenta con señalización en ambas áreas.",
                "No aplica.",
            ],
            "options_secondary": [
                "La señalización cumple con los criterios de altura de acuerdo a la NTC (Altura de 120 cm a 160 cm).",
                "La señalización No cumple con los criterios de altura de acuerdo a la NTC (Altura de 120 cm a 160 cm).",
                "No aplica.",
                "No se cuenta con señalización de orientación.",
            ],
            "options_tertiary": [
                "La señalización cumple con los criterios de material de acuerdo a la NTC.",
                "La señalización No cumple con los criterios de material de acuerdo a la NTC.",
                "No aplica.",
            ],
        },
        {
            "id": "ajustes_razonables_individualizados",
            "label": "¿La organización cuenta con la posibilidad de hacer ajustes razonables individualizados?",
            "type": "lista",
            "has_accesible": True,
            "has_observaciones": False,
            "options": [
                "Cuenta con la posibilidad de flexibilizar rutinas laborales.",
                "No cuenta con la posibilidad de flexibilizar rutinas laborales.",
                "No aplica.",
            ],
        },
        {
            "id": "ajustes_razonables_detalle",
            "label": "Detalle de ajustes razonables individualizados",
            "type": "texto",
            "has_accesible": False,
            "has_observaciones": False,
        },
    ],
    "accesible_options": ["Sí", "No", "Parcial"],
}


SECTION_2_4 = seccion_2_4.SECTION_2_4
SECTION_2_5 = seccion_2_5_2_6.SECTION_2_5
SECTION_2_6 = seccion_2_5_2_6.SECTION_2_6
SECTION_3 = seccion_3.SECTION_3
SECTION_4 = seccion_4.SECTION_4
SECTION_5 = seccion_5.SECTION_5
SECTION_6 = seccion_6_7.SECTION_6
SECTION_7 = seccion_6_7.SECTION_7
SECTION_8 = seccion_8.SECTION_8

def get_empresa_by_nit(nit, env_path=".env"):
    if not nit:
        return None
    nit = "".join(str(nit).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))

    def _normalize_nit(value):
        return re.sub(r"[^0-9A-Za-z]+", "", str(value or "")).lower()

    def _nit_candidates(value):
        raw = "".join(str(value or "").split())
        compact = re.sub(r"[^0-9A-Za-z]+", "", raw)
        candidates = []
        for cand in (raw, compact):
            if cand and cand not in candidates:
                candidates.append(cand)
        if compact.isdigit() and len(compact) > 1:
            with_dash = f"{compact[:-1]}-{compact[-1]}"
            if with_dash not in candidates:
                candidates.append(with_dash)
        return candidates

    # 1) Intentos exactos con variantes comunes.
    for candidate in _nit_candidates(nit):
        params = {
            "select": select_cols,
            "nit_empresa": f"eq.{candidate}",
            "limit": 1,
        }
        data = _supabase_get("empresas", params, env_path=env_path)
        if data:
            return data[0]

    # 2) Fallback por coincidencia normalizada (con/sin guion/simbolos).
    normalized_target = _normalize_nit(nit)
    if not normalized_target:
        return None
    params = {
        "select": select_cols,
        "nit_empresa": f"ilike.%{normalized_target}%",
        "limit": 20,
    }
    rows = _supabase_get("empresas", params, env_path=env_path)
    if not rows:
        return None
    exact = [row for row in rows if _normalize_nit(row.get("nit_empresa")) == normalized_target]
    if len(exact) == 1:
        return exact[0]
    if len(rows) == 1:
        return rows[0]
    return None


def get_empresa_by_nombre(nombre, env_path=".env"):
    if not nombre:
        return None
    nombre = " ".join(str(nombre).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nombre_empresa": f"ilike.{nombre}",
        "limit": 5,
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    if len(data) == 1:
        return data[0]

    def _normalize_name(value):
        return re.sub(r"\s+", " ", str(value or "")).strip().lower()

    target = _normalize_name(nombre)
    # Fallback por coincidencia parcial para bases grandes y variaciones menores.
    params = {
        "select": select_cols,
        "nombre_empresa": f"ilike.%{nombre}%",
        "limit": 50,
    }
    candidates = _supabase_get("empresas", params, env_path=env_path)
    if not candidates:
        return None
    exact = [row for row in candidates if _normalize_name(row.get("nombre_empresa")) == target]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        raise ValueError("Hay más de una empresa con ese nombre. Usa el NIT.")
    if len(candidates) == 1:
        return candidates[0]
    raise ValueError("Hay más de una empresa con ese nombre. Usa el NIT.")


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=10):
    if not prefix:
        return []
    prefix = " ".join(str(prefix).split())
    if not prefix:
        return []
    params = {
        "select": "nombre_empresa",
        "nombre_empresa": f"ilike.{prefix}%",
        "limit": max(1, int(limit)),
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    if not data:
        return []
    names = []
    seen = set()
    for row in data:
        name = row.get("nombre_empresa")
        if not name:
            continue
        if name in seen:
            continue
        seen.add(name)
        names.append(name)
    return names


def confirm_section_1(company_data, user_inputs):
    if not company_data:
        raise ValueError("No hay datos de empresa para confirmar.")
    payload = {}
    for field in SECTION_1["fields"]:
        field_id = field["id"]
        if field["source"] == "input":
            payload[field_id] = user_inputs.get(field_id)
        else:
            payload[field_id] = company_data.get(field_id)
    SECTION_1_CACHE.update(payload)
    set_section_cache("section_1", payload)
    FORM_CACHE["_last_section"] = "section_1"
    save_cache_to_file()
    return payload


def confirm_section_2_1(payload):
    if payload is None:
        raise ValueError("section_2_1 requerida")
    set_section_cache("section_2_1", payload)
    FORM_CACHE["_last_section"] = "section_2_1"
    save_cache_to_file()
    return payload


def confirm_section_2_2(payload):
    if payload is None:
        raise ValueError("section_2_2 requerida")
    set_section_cache("section_2_2", payload)
    FORM_CACHE["_last_section"] = "section_2_2"
    save_cache_to_file()
    return payload


def confirm_section_2_3(payload):
    if payload is None:
        raise ValueError("section_2_3 requerida")
    set_section_cache("section_2_3", payload)
    FORM_CACHE["_last_section"] = "section_2_3"
    save_cache_to_file()
    return payload


def confirm_section_2_4(payload):
    if payload is None:
        raise ValueError("section_2_4 requerida")
    set_section_cache("section_2_4", payload)
    FORM_CACHE["_last_section"] = "section_2_4"
    save_cache_to_file()
    return payload


def confirm_section_2_5(payload):
    if payload is None:
        raise ValueError("section_2_5 requerida")
    set_section_cache("section_2_5", payload)
    FORM_CACHE["_last_section"] = "section_2_5"
    save_cache_to_file()
    return payload


def confirm_section_2_6(payload):
    if payload is None:
        raise ValueError("section_2_6 requerida")
    set_section_cache("section_2_6", payload)
    FORM_CACHE["_last_section"] = "section_2_6"
    save_cache_to_file()
    return payload


def confirm_section_3(payload):
    if payload is None:
        raise ValueError("section_3 requerida")
    set_section_cache("section_3", payload)
    FORM_CACHE["_last_section"] = "section_3"
    save_cache_to_file()
    return payload


def confirm_section_4(payload):
    if payload is None:
        raise ValueError("section_4 requerida")
    set_section_cache("section_4", payload)
    FORM_CACHE["_last_section"] = "section_4"
    save_cache_to_file()
    return payload


def confirm_section_5(payload):
    if payload is None:
        raise ValueError("section_5 requerida")
    for item in SECTION_5.get("items", []):
        field_id = item.get("id")
        if not field_id:
            continue
        aplica_value = payload.get(field_id)
        if aplica_value == "Aplica":
            payload[f"{field_id}_ajustes"] = item.get("ajustes", "")
        else:
            payload[f"{field_id}_ajustes"] = "No aplica"
    set_section_cache("section_5", payload)
    FORM_CACHE["_last_section"] = "section_5"
    save_cache_to_file()
    return payload


def confirm_section_6(payload):
    if payload is None:
        raise ValueError("section_6 requerida")
    set_section_cache("section_6", payload)
    FORM_CACHE["_last_section"] = "section_6"
    save_cache_to_file()
    return payload


def confirm_section_7(payload):
    if payload is None:
        raise ValueError("section_7 requerida")
    set_section_cache("section_7", payload)
    FORM_CACHE["_last_section"] = "section_7"
    save_cache_to_file()
    return payload


def confirm_section_8(payload):
    if payload is None:
        raise ValueError("section_8 requerida")
    set_section_cache("section_8", payload)
    FORM_CACHE["_last_section"] = "section_8"
    save_cache_to_file()
    return payload


def set_section_cache(section_id, payload):
    if not section_id:
        raise ValueError("section_id requerido")
    if payload is None:
        payload = {}
    FORM_CACHE[section_id] = payload


def get_form_cache():
    return dict(FORM_CACHE)


def register_form():
    """Return metadata for hub registration."""
    return {
        "id": "evaluacion_accesibilidad",
        "name": FORM_NAME,
        "module": __name__,
    }

