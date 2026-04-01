#!/usr/bin/env python
"""
E2E — Contratacion Incluyente
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

from formularios.contratacion_incluyente import contratacion_incluyente as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "contratacion_incluyente"

# Primera opción de cada dropdown
_DISCAPACIDAD = "Discapacidad física"
_GENERO = "Binario"
_LGTBIQ = "No aplica"
_GRUPO_ETNICO = "No aplica"
_GRUPO_ETNICO_CUAL = "No aplica"
_CERT_DISC = "No"
_TIPO_CONTRATO = "Contrato de trabajo a término fijo"
_NIVEL_APOYO = "0. No requiere apoyo."
_JORNADA = "Tiempo Completo."
_CLAUSULAS = "Cláusula de confidencialidad."
_FRECUENCIA_PAGO = "Pago Quincenal."
_FORMA_PAGO = "Abono a cuenta bancaria."


def _build_vinculado():
    """Un vinculado con todos los campos necesarios para el export."""
    return {
        "numero": "1",
        "nombre_oferente": TEXTO_CORTO,
        "cedula": TEXTO_CORTO,
        "certificado_porcentaje": TEXTO_CORTO,
        "discapacidad": _DISCAPACIDAD,
        "telefono_oferente": TEXTO_CORTO,
        "genero": _GENERO,
        "correo_oferente": TEXTO_CORTO,
        "fecha_nacimiento": TEXTO_CORTO,
        "edad": TEXTO_CORTO,
        "lgtbiq": _LGTBIQ,
        "grupo_etnico": _GRUPO_ETNICO,
        "grupo_etnico_cual": _GRUPO_ETNICO_CUAL,
        "cargo_oferente": TEXTO_CORTO,
        "contacto_emergencia": TEXTO_CORTO,
        "parentesco": TEXTO_CORTO,
        "telefono_emergencia": TEXTO_CORTO,
        "certificado_discapacidad": _CERT_DISC,
        "lugar_firma_contrato": TEXTO_CORTO,
        "fecha_firma_contrato": TEXTO_CORTO,
        "tipo_contrato": _TIPO_CONTRATO,
        "fecha_fin": TEXTO_CORTO,
        # Sección 5.1 - Lectura del contrato
        "contrato_lee_nivel_apoyo": _NIVEL_APOYO,
        "contrato_lee_observacion": TEXTO_CORTO,
        "contrato_lee_nota": TEXTO_CORTO,
        "contrato_comprendido_nivel_apoyo": _NIVEL_APOYO,
        "contrato_comprendido_observacion": TEXTO_CORTO,
        "contrato_comprendido_nota": TEXTO_CORTO,
        "contrato_tipo_nivel_apoyo": _NIVEL_APOYO,
        "contrato_tipo_observacion": TEXTO_CORTO,
        "contrato_tipo_contrato": _TIPO_CONTRATO,
        "contrato_jornada": _JORNADA,
        "contrato_clausulas": _CLAUSULAS,
        "contrato_tipo_nota": TEXTO_CORTO,
        # Condiciones salariales
        "condiciones_salariales_nivel_apoyo": _NIVEL_APOYO,
        "condiciones_salariales_observacion": TEXTO_CORTO,
        "condiciones_salariales_frecuencia_pago": _FRECUENCIA_PAGO,
        "condiciones_salariales_forma_pago": _FORMA_PAGO,
        "condiciones_salariales_nota": TEXTO_CORTO,
        # Sección 5.2 - Prestaciones de ley
        "prestaciones_cesantias_nivel_apoyo": _NIVEL_APOYO,
        "prestaciones_cesantias_observacion": TEXTO_CORTO,
        "prestaciones_cesantias_nota": TEXTO_CORTO,
        "prestaciones_auxilio_transporte_nivel_apoyo": _NIVEL_APOYO,
        "prestaciones_auxilio_transporte_observacion": TEXTO_CORTO,
        "prestaciones_auxilio_transporte_nota": TEXTO_CORTO,
        "prestaciones_prima_nivel_apoyo": _NIVEL_APOYO,
        "prestaciones_prima_observacion": TEXTO_CORTO,
        "prestaciones_prima_nota": TEXTO_CORTO,
        "prestaciones_seguridad_social_nivel_apoyo": _NIVEL_APOYO,
        "prestaciones_seguridad_social_observacion": TEXTO_CORTO,
        "prestaciones_seguridad_social_nota": TEXTO_CORTO,
        "prestaciones_vacaciones_nivel_apoyo": _NIVEL_APOYO,
        "prestaciones_vacaciones_observacion": TEXTO_CORTO,
        "prestaciones_vacaciones_nota": TEXTO_CORTO,
        "prestaciones_auxilios_beneficios_nivel_apoyo": _NIVEL_APOYO,
        "prestaciones_auxilios_beneficios_observacion": TEXTO_CORTO,
        "prestaciones_auxilios_beneficios_nota": TEXTO_CORTO,
        # Sección 5.3 - Deberes y derechos
        "conducto_regular_nivel_apoyo": _NIVEL_APOYO,
        "conducto_regular_observacion": TEXTO_CORTO,
        "descargos_observacion": TEXTO_CORTO,
        "tramites_observacion": TEXTO_CORTO,
        "permisos_observacion": TEXTO_CORTO,
        "conducto_regular_nota": TEXTO_CORTO,
        "causales_fin_nivel_apoyo": _NIVEL_APOYO,
        "causales_fin_observacion": TEXTO_CORTO,
        "causales_fin_nota": TEXTO_CORTO,
        "rutas_atencion_nivel_apoyo": _NIVEL_APOYO,
        "rutas_atencion_observacion": TEXTO_CORTO,
        "rutas_atencion_nota": TEXTO_CORTO,
        # Desarrollo de la actividad
        "desarrollo_actividad": TEXTO_LARGO,
    }


def run():
    # --- Limpiar cache ---
    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    # --- Sección 1: Datos generales ---
    s1 = dict(SECTION_1_BASE)
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # --- Sección 2: Vinculados ---
    mod.FORM_CACHE["section_2"] = [_build_vinculado()]

    # --- Sección 6: Ajustes razonables ---
    mod.FORM_CACHE["section_6"] = {"ajustes_recomendaciones": TEXTO_LARGO}

    # --- Sección 7: Asistentes ---
    mod.FORM_CACHE["section_7"] = list(ASISTENTES)

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
