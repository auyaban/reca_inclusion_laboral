#!/usr/bin/env python
"""
E2E — Induccion Operativa
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

from formularios.induccion_operativa import induccion_operativa as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "induccion_operativa"

_EJECUCION = "Si"
_NIVEL_APOYO_S4 = "0. No requiere apoyo."
_NIVEL_APOYO_S5 = "No requiero apoyo."

# IDs de items de sección 3
_SECTION_3_ITEMS = [
    "funciones_corresponden_perfil",
    "explicacion_funciones",
    "instrucciones_claras",
    "sistema_medicion",
    "induccion_maquinas",
    "presentacion_companeros",
    "presentacion_jefes",
    "uso_epp",
    "conducto_regular",
    "puesto_trabajo",
    "otros",
]

# Bloques de sección 4
_SECTION_4_BLOCKS = [
    "comprension_instrucciones",
    "autonomia_tareas",
    "trabajo_equipo",
    "adaptacion_flexibilidad",
    "solucion_problemas",
    "comunicacion_asertiva",
    "manejo_tiempo",
    "iniciativa_proactividad",
]

# Items de sección 5
_SECTION_5_ITEMS = [
    "condiciones_medicas_salud",
    "habilidades_basicas_vida_diaria",
    "habilidades_socioemocionales",
]


def run():
    # --- Limpiar cache ---
    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    # --- Sección 1: Datos generales ---
    s1 = dict(SECTION_1_BASE)
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # --- Sección 2: Vinculados ---
    mod.FORM_CACHE["section_2"] = [
        {
            "numero": "1",
            "nombre_oferente": TEXTO_CORTO,
            "cedula": TEXTO_CORTO,
            "telefono_oferente": TEXTO_CORTO,
            "cargo_oferente": TEXTO_CORTO,
        }
    ]

    # --- Sección 3: Ejecución del proceso ---
    s3 = {}
    for item_id in _SECTION_3_ITEMS:
        s3[item_id] = {
            "ejecucion": _EJECUCION,
            "observaciones": TEXTO_CORTO,
        }
    mod.FORM_CACHE["section_3"] = s3

    # --- Sección 4: Habilidades socioemocionales ---
    # items: cada item individual con nivel_apoyo y observaciones
    # notes: nota de resumen por bloque
    _SECTION_4_ITEMS = [
        "reconoce_instrucciones",
        "proceso_atencion",
        "identifica_funciones",
        "importancia_calidad",
        "relacion_companeros",
        "recibe_sugerencias",
        "objetivos_grupales",
        "reconoce_entorno",
        "ajuste_cambios",
        "identifica_problema_laboral",
        "respeto_companeros",
        "lenguaje_corporal",
        "reporte_novedades",
        "organiza_actividades",
        "cumple_horario",
        "identifica_horarios",
        "reporta_finalizacion",
    ]
    s4_items = {item_id: {"nivel_apoyo": _NIVEL_APOYO_S4, "observaciones": TEXTO_CORTO}
                for item_id in _SECTION_4_ITEMS}
    s4_notes = {block_id: TEXTO_CORTO for block_id in _SECTION_4_BLOCKS}
    mod.FORM_CACHE["section_4"] = {"items": s4_items, "notes": s4_notes}

    # --- Sección 5: Nivel de apoyo requerido ---
    s5 = {}
    for item_id in _SECTION_5_ITEMS:
        s5[item_id] = {
            "nivel_apoyo_requerido": _NIVEL_APOYO_S5,
            "observaciones": TEXTO_CORTO,
        }
    mod.FORM_CACHE["section_5"] = s5

    # --- Sección 6: Ajustes razonables ---
    mod.FORM_CACHE["section_6"] = {"ajustes_requeridos": TEXTO_LARGO}

    # --- Sección 7: Fecha primer seguimiento ---
    mod.FORM_CACHE["section_7"] = {"fecha_primer_seguimiento": "2026-06-05"}

    # --- Sección 8: Observaciones y recomendaciones ---
    mod.FORM_CACHE["section_8"] = {"observaciones_recomendaciones": TEXTO_LARGO}

    # --- Sección 9: Asistentes ---
    mod.FORM_CACHE["section_9"] = list(ASISTENTES)

    # --- Medir tiempo de finalización ---
    print(f"[{FORM_ID}] Iniciando finalización...")
    t0 = time.time()
    try:
        result = mod.export_to_excel(clear_cache=False)
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
