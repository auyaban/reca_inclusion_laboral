#!/usr/bin/env python
"""
E2E — Seguimientos
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

from formularios.seguimientos import seguimientos as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "seguimientos"

# Cedula ficticia para el E2E (no colisionar con casos reales)
_CEDULA = "99999999"
_NOMBRE_USUARIO = "Pendiente E2E"

# user_row describe al vinculado + empresa para la hoja de seguimiento
_USER_ROW = {
    "nombre_usuario": _NOMBRE_USUARIO,
    "cedula_usuario": _CEDULA,
    "telefono_oferente": TEXTO_CORTO,
    "correo_oferente": TEXTO_CORTO,
    "contacto_emergencia": TEXTO_CORTO,
    "parentesco": TEXTO_CORTO,
    "telefono_emergencia": TEXTO_CORTO,
    "cargo_oferente": TEXTO_CORTO,
    "certificado_discapacidad": "No",
    "certificado_porcentaje": TEXTO_CORTO,
    "discapacidad_usuario": "Discapacidad fisica",
    "tipo_contrato": "Contrato de trabajo a termino fijo",
    "fecha_firma_contrato": SECTION_1_BASE["fecha_visita"],
    # Datos empresa (fallback si Supabase no responde)
    "empresa_nit": SECTION_1_BASE["nit_empresa"],
    "empresa_nombre": SECTION_1_BASE["nombre_empresa"],
}

_FOLLOWUP_PAYLOAD = {
    "modalidad": SECTION_1_BASE["modalidad"],
    "seguimiento_numero": 1,
    "tipo_apoyo": "No requiere apoyo.",
    "situacion_encontrada": TEXTO_LARGO,
    "estrategias_ajustes": TEXTO_LARGO,
    "item_observaciones": [TEXTO_CORTO] * 19,
    "item_autoevaluacion": ["Excelente"] * 19,
    "item_eval_empresa": ["Excelente"] * 19,
    "empresa_eval": ["Excelente"] * 8,
    "empresa_observacion": [TEXTO_CORTO] * 8,
    "asistentes": list(ASISTENTES),
}


def run():
    print(f"[{FORM_ID}] Iniciando creacion de caso + seguimiento...")
    t0 = time.time()
    try:
        result = mod.ensure_case_record(_CEDULA, _USER_ROW, is_compensar=True)
        record = result["record"]

        # Escribir el primer seguimiento
        mod.save_followup_payload(record, 1, _FOLLOWUP_PAYLOAD)

    except Exception as exc:
        elapsed = time.time() - t0
        print(f"[ERROR] [{FORM_ID}] fallo en {elapsed:.2f}s: {exc}")
        traceback.print_exc()
        return {
            "form_id": FORM_ID,
            "status": "ERROR",
            "elapsed_s": round(elapsed, 2),
            "url": None,
            "error": str(exc),
        }
    elapsed = time.time() - t0

    url = record.get("webViewLink") or ""
    status = "OK" if url else "NO_URL"
    created = result.get("created", False)
    print(f"[{status}] [{FORM_ID}] {elapsed:.2f}s -> {url} (created={created})")
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
