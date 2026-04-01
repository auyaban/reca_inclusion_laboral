#!/usr/bin/env python
"""
E2E — Reactivacion del Programa IL
Empresa: CORONA INDUSTRIAL SAS (NIT 900696296-4)
Fecha: 2026-05-05, Modalidad: Presencial
"""
import sys
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

from formularios.presentacion_programa import presentacion_programa as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "reactivacion_programa"


def run():
    # --- Limpiar cache ---
    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    # --- Sección 1: tipo_visita = Reactivación ---
    s1 = dict(SECTION_1_BASE)
    s1["tipo_visita"] = "Reactivación"
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # --- Sección 3 / item 8: Motivación empresarial (checkboxes) ---
    mod.FORM_CACHE["section_3_item_8"] = {
        "Responsabilidad Social Empresarial": True,
        "Objetivos y metas para la diversidad, equidad e inclusión": True,
        "Avances a nivel global de impacto en Colombia": True,
        "Beneficios Tributarios": True,
        "Beneficios en la contratación de población en riesgo de exclusión": True,
        "Ventaja en licitaciones públicas": True,
        "Cumplimiento de la normativa establecida por el Estado Colombiano": True,
        "Experiencia en la vinculación de personas en condición de discapacidad": True,
    }

    # --- Sección 4: Acuerdos y observaciones ---
    mod.FORM_CACHE["section_4"] = {"acuerdos_observaciones": TEXTO_LARGO}
    mod.FORM_CACHE["section_3"] = {"observaciones": TEXTO_LARGO}

    # --- Sección 5: Asistentes ---
    mod.FORM_CACHE["section_5"] = list(ASISTENTES)

    # --- Medir tiempo de finalización ---
    print(f"[{FORM_ID}] Iniciando finalización...")
    t0 = time.time()
    try:
        result = mod.export_to_excel(cache=mod.FORM_CACHE)
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
