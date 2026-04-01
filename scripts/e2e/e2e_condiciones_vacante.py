#!/usr/bin/env python
"""
E2E — Condiciones de Vacante
Empresa: CORONA INDUSTRIAL SAS (NIT 900696296-4)
Fecha: 2026-05-05, Modalidad: Presencial
"""
import sys
import io
import time
import traceback
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from formularios.condiciones_vacante import condiciones_vacante as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "condiciones_vacante"
MODALIDAD = "Presencial"

# Valores representativos de primera opción de cada dropdown
_NIVEL_CARGO = "Administrativo."
_GENERO = "Hombre - Mujer"
_TIPO_CONTRATO = "Contrato de trabajo a término fijo"
_REQ_CERT = "No"
_HABILIDAD = "Alto."
_NO_APLICA = "No aplica"
_TIEMPO = "De 6 a 8 horas."
_FRECUENCIA = "Diario."
_RIESGO = "Bajo."


def run():
    # --- Limpiar cache ---
    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    # --- Sección 1: Datos generales ---
    s1 = dict(SECTION_1_BASE)
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # --- Sección 2: Características de la vacante ---
    mod.FORM_CACHE["section_2"] = {
        "nombre_vacante": TEXTO_CORTO,
        "numero_vacantes": "1",
        "nivel_cargo": _NIVEL_CARGO,
        "genero": _GENERO,
        "edad": TEXTO_CORTO,
        "modalidad_trabajo": MODALIDAD,
        "lugar_trabajo": TEXTO_CORTO,
        "salario_asignado": TEXTO_CORTO,
        "firma_contrato": TEXTO_CORTO,
        "aplicacion_pruebas": TEXTO_CORTO,
        "tipo_contrato": _TIPO_CONTRATO,
        "beneficios_adicionales": TEXTO_CORTO,
        "cargo_flexible_genero": TEXTO_CORTO,
        "beneficios_mujeres": TEXTO_CORTO,
        "requiere_certificado": _REQ_CERT,
        "requiere_certificado_observaciones": TEXTO_CORTO,
        "competencia_1": TEXTO_CORTO,
        "competencia_2": TEXTO_CORTO,
        "competencia_3": TEXTO_CORTO,
        "competencia_4": TEXTO_CORTO,
        "competencia_5": TEXTO_CORTO,
        "competencia_6": TEXTO_CORTO,
        "competencia_7": TEXTO_CORTO,
        "competencia_8": TEXTO_CORTO,
    }

    # --- Sección 2.1: Formación y experiencia ---
    mod.FORM_CACHE["section_2_1"] = {
        "nivel_bachiller": True,
        "nivel_tecnico_profesional": True,
        "nivel_primaria": False,
        "nivel_profesional": False,
        "nivel_especializacion": False,
        "nivel_tecnologo": False,
        "especificaciones_formacion": TEXTO_LARGO,
        "conocimientos_basicos": TEXTO_LARGO,
        "horarios_asignados": TEXTO_CORTO,
        "hora_ingreso": "7:00 a.m.",
        "hora_salida": "5:00 p.m.",
        "tiempo_almuerzo": "1 hora.",
        "break_descanso": "15 minutos",
        "dias_laborables": "Lunes a Viernes",
        "dias_flexibles": _NO_APLICA,
        "observaciones": TEXTO_LARGO,
        "experiencia_meses": TEXTO_CORTO,
        "funciones_tareas": TEXTO_LARGO,
        "herramientas_equipos": TEXTO_LARGO,
    }

    # --- Sección 3: Habilidades ---
    habilidades_fields = [
        "lectura", "comprension_lectora", "escritura", "comunicacion_verbal",
        "razonamiento_logico", "conteo_reporte", "clasificacion_objetos",
        "velocidad_ejecucion", "concentracion", "memoria", "ubicacion_espacial",
        "atencion", "agarre", "precision", "digitacion", "agilidad_manual",
        "coordinacion_ojo_mano", "esfuerzo_fisico", "equilibrio_corporal",
        "lanzar_objetos", "seguimiento_instrucciones", "resolucion_conflictos",
        "autonomia_tareas", "trabajo_equipo", "adaptabilidad", "flexibilidad",
        "comunicacion_asertiva", "manejo_tiempo", "liderazgo", "escucha_activa",
        "proactividad",
    ]
    s3 = {f: _HABILIDAD for f in habilidades_fields}
    s3["observaciones_cognitivas"] = TEXTO_LARGO
    s3["observaciones_motricidad_fina"] = TEXTO_LARGO
    s3["observaciones_motricidad_gruesa"] = TEXTO_LARGO
    s3["observaciones_transversales"] = TEXTO_LARGO
    mod.FORM_CACHE["section_3"] = s3

    # --- Sección 4: Posturas (filas 111-115) ---
    mod.FORM_CACHE["section_4"] = {
        "sentado_tiempo": _TIEMPO,
        "sentado_frecuencia": _FRECUENCIA,
        "semisentado_tiempo": _NO_APLICA,
        "semisentado_frecuencia": _NO_APLICA,
        "de_pie_tiempo": "De 1 a 2 horas.",
        "de_pie_frecuencia": _FRECUENCIA,
        "agachado_tiempo": _NO_APLICA,
        "agachado_frecuencia": _NO_APLICA,
        "uso_extremidades_superiores_tiempo": _TIEMPO,
        "uso_extremidades_superiores_frecuencia": _FRECUENCIA,
    }

    # --- Sección 5: Peligros (filas 120-150) ---
    mod.FORM_CACHE["section_5"] = {
        "ruido": _RIESGO,
        "iluminacion": "Medio.",
        "temperaturas_externas": _RIESGO,
        "vibraciones": _NO_APLICA,
        "presion_atmosferica": _NO_APLICA,
        "radiaciones": _RIESGO,
        "polvos_organicos_inorganicos": _NO_APLICA,
        "fibras": _NO_APLICA,
        "liquidos": _NO_APLICA,
        "gases_vapores": _NO_APLICA,
        "humos_metalicos": _NO_APLICA,
        "humos_no_metalicos": _NO_APLICA,
        "material_particulado": _NO_APLICA,
        "electrico": _RIESGO,
        "locativo": _RIESGO,
        "accidentes_transito": _NO_APLICA,
        "publicos": _NO_APLICA,
        "mecanico": _NO_APLICA,
        "gestion_organizacional": "Medio.",
        "caracteristicas_organizacion": _RIESGO,
        "caracteristicas_grupo_social": _RIESGO,
        "condiciones_tarea": "Medio.",
        "interfase_persona_tarea": "Medio.",
        "jornada_trabajo": _RIESGO,
        "postura_trabajo": "Medio.",
        "puesto_trabajo": _RIESGO,
        "movimientos_repetitivos": "Alto.",
        "manipulacion_cargas": _NO_APLICA,
        "herramientas_equipos": _RIESGO,
        "organizacion_trabajo": _RIESGO,
        "observaciones_peligros": TEXTO_LARGO,
    }

    # --- Sección 6: Discapacidades identificadas ---
    mod.FORM_CACHE["section_6"] = [
        {
            "discapacidad": "DISCAPACIDAD FÍSICA",
            "descripcion": TEXTO_LARGO,
        }
    ]

    # --- Sección 7: Observaciones y recomendaciones ---
    mod.FORM_CACHE["section_7"] = {"observaciones_recomendaciones": TEXTO_LARGO}

    # --- Sección 8: Asistentes ---
    mod.FORM_CACHE["section_8"] = list(ASISTENTES)

    # --- Medir tiempo de finalización ---
    print(f"[{FORM_ID}] Iniciando finalización...")
    t0 = time.time()
    try:
        result = mod.export_to_excel(progress_callback=None)
    except Exception as exc:
        elapsed = time.time() - t0
        print(f"[ERROR] [{FORM_ID}] export falló en {elapsed:.2f}s: {exc}")
        traceback.print_exc()
        return {
            "form_id": FORM_ID,
            "status": "ERROR",
            "elapsed_s": round(elapsed, 2),
            "url": None,
            "error": str(exc),
        }
    elapsed = time.time() - t0

    url = result.get("output_path", "") if isinstance(result, dict) else str(result or "")
    status = "OK" if url else "NO_URL"
    print(f"[{status}] [{FORM_ID}] {elapsed:.2f}s -> {url}")
    return {
        "form_id": FORM_ID,
        "status": status,
        "elapsed_s": round(elapsed, 2),
        "url": url,
        "error": None,
    }


if __name__ == "__main__":
    r = run()
    sys.exit(0 if r["status"] == "OK" else 1)
