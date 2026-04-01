#!/usr/bin/env python
"""
E2E — Induccion Organizacional
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

from formularios.induccion_organizacional import induccion_organizacional as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "induccion_organizacional"

_VISTO = "Si"
_MEDIO = "Mixto"
_MEDIO_SEC4 = "Video"


# Todos los items de sección 3 (subsecciones 3.1 – 3.5)
_SECTION_3_ITEMS = [
    # 3.1 Generalidades
    "historia_empresa", "mision_organizacional", "vision_organizacional",
    "objetivos_valores_principios", "recorrido_empresa",
    # 3.2 Gestión Humana
    "tramites_permisos", "formas_pago", "obligaciones_prohibiciones",
    "normatividad_interna", "practicas_inclusivas", "horario_laboral",
    "organigrama", "incapacidades_permisos_calamidades", "equipos_tecnologicos",
    "comites", "conductos_regulares_comunicacion",
    # 3.3 SG-SST
    "sgsst_general", "peligros_riesgos", "uso_epp", "politicas_medio_ambiente",
    "politicas_confidencialidad", "plan_emergencias", "prevencion_consumo",
    "normas_comite", "normas_disciplinarias", "entrega_dotacion_epp",
    "brigada_emergencia", "mecanismos_desempeno", "procedimiento_accidente",
    # 3.4 Inducción al puesto
    "funciones_especificas", "horario_turnos", "dotacion_uniformes",
    "presentacion_equipo", "registro_ingreso", "entrega_carnet", "recorrido_puesto",
    # 3.5 Evaluativo
    "evaluaciones", "plataformas_elearning",
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

    # --- Sección 3: Desarrollo del proceso ---
    s3 = {}
    for item_id in _SECTION_3_ITEMS:
        s3[item_id] = {
            "visto": _VISTO,
            "responsable": TEXTO_CORTO,
            "medio_socializacion": _MEDIO,
            "descripcion": TEXTO_CORTO,
        }
    mod.FORM_CACHE["section_3"] = s3

    # --- Sección 4: Recomendaciones de accesibilidad (lista de 3) ---
    mod.FORM_CACHE["section_4"] = [
        {"medio": _MEDIO_SEC4},
        {"medio": _MEDIO_SEC4},
        {"medio": _MEDIO_SEC4},
    ]

    # --- Sección 5: Observaciones ---
    mod.FORM_CACHE["section_5"] = {"observaciones": TEXTO_LARGO}

    # --- Sección 6: Asistentes ---
    mod.FORM_CACHE["section_6"] = list(ASISTENTES)

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
