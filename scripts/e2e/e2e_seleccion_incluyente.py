#!/usr/bin/env python
"""
E2E — Proceso de Seleccion Incluyente (template individual)
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

from formularios.seleccion_incluyente import seleccion_incluyente as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "seleccion_incluyente"

# Primera opción de cada dropdown
_DISCAPACIDAD = "Discapacidad física"
_NIVEL_APOYO = "0. No requiere apoyo."
_PENSION = "No aplica"
_TIPO_PENSION = "No aplica"
_PENDIENTE = "Por Confirmar"
_DINERO = "Autonomo."


def _build_oferente():
    """Un oferente con todos los campos mínimos para el export."""
    return {
        "numero": "1",
        "nombre_oferente": TEXTO_CORTO,
        "cedula": TEXTO_CORTO,
        "certificado_porcentaje": TEXTO_CORTO,
        "discapacidad": _DISCAPACIDAD,
        "telefono_oferente": TEXTO_CORTO,
        "resultado_certificado": "Pendiente",
        "cargo_oferente": TEXTO_CORTO,
        "nombre_contacto_emergencia": TEXTO_CORTO,
        "parentesco": TEXTO_CORTO,
        "telefono_emergencia": TEXTO_CORTO,
        "fecha_nacimiento": TEXTO_CORTO,
        "edad": TEXTO_CORTO,
        "pendiente_otros_oferentes": _PENDIENTE,
        "lugar_firma_contrato": TEXTO_CORTO,
        "fecha_firma_contrato": TEXTO_CORTO,
        "cuenta_pension": _PENDIENTE,
        "tipo_pension": _TIPO_PENSION,
        "desarrollo_actividad": TEXTO_LARGO,
        # Medicamentos
        "medicamentos_nivel_apoyo": _NIVEL_APOYO,
        "medicamentos_conocimiento": TEXTO_CORTO,
        "medicamentos_horarios": TEXTO_CORTO,
        "medicamentos_nota": TEXTO_CORTO,
        # Alergias
        "alergias_nivel_apoyo": _NIVEL_APOYO,
        "alergias_tipo": TEXTO_CORTO,
        "alergias_nota": TEXTO_CORTO,
        # Restricciones
        "restriccion_nivel_apoyo": _NIVEL_APOYO,
        "restriccion_conocimiento": TEXTO_CORTO,
        "restriccion_nota": TEXTO_CORTO,
        # Controles médicos
        "controles_nivel_apoyo": _NIVEL_APOYO,
        "controles_asistencia": TEXTO_CORTO,
        "controles_frecuencia": TEXTO_CORTO,
        "controles_nota": TEXTO_CORTO,
        # Desplazamiento
        "desplazamiento_nivel_apoyo": _NIVEL_APOYO,
        "desplazamiento_modo": TEXTO_CORTO,
        "desplazamiento_transporte": TEXTO_CORTO,
        "desplazamiento_nota": TEXTO_CORTO,
        # Ubicación
        "ubicacion_nivel_apoyo": _NIVEL_APOYO,
        "ubicacion_ciudad": TEXTO_CORTO,
        "ubicacion_aplicaciones": TEXTO_CORTO,
        "ubicacion_nota": TEXTO_CORTO,
        # Dinero
        "dinero_nivel_apoyo": _NIVEL_APOYO,
        "dinero_reconocimiento": _DINERO,
        "dinero_manejo": TEXTO_CORTO,
        "dinero_medios": TEXTO_CORTO,
        "dinero_nota": TEXTO_CORTO,
        # Presentación personal
        "presentacion_nivel_apoyo": _NIVEL_APOYO,
        "presentacion_personal": TEXTO_CORTO,
        "presentacion_nota": TEXTO_CORTO,
        # Comunicación escrita
        "comunicacion_escrita_nivel_apoyo": _NIVEL_APOYO,
        "comunicacion_escrita_apoyo": TEXTO_CORTO,
        "comunicacion_escrita_nota": TEXTO_CORTO,
        # Comunicación verbal
        "comunicacion_verbal_nivel_apoyo": _NIVEL_APOYO,
        "comunicacion_verbal_apoyo": TEXTO_CORTO,
        "comunicacion_verbal_nota": TEXTO_CORTO,
        # Toma de decisiones
        "decisiones_nivel_apoyo": _NIVEL_APOYO,
        "toma_decisiones": TEXTO_CORTO,
        "toma_decisiones_nota": TEXTO_CORTO,
        # Aseo y alimentación
        "aseo_nivel_apoyo": _NIVEL_APOYO,
        "alimentacion": TEXTO_CORTO,
        "aseo_criar_apoyo": TEXTO_CORTO,
        "aseo_comunicacion_apoyo": TEXTO_CORTO,
        "aseo_ayudas_apoyo": TEXTO_CORTO,
        "aseo_alimentacion": TEXTO_CORTO,
        "aseo_movilidad_funcional": TEXTO_CORTO,
        "aseo_higiene_aseo": TEXTO_CORTO,
        "aseo_nota": TEXTO_CORTO,
        # Actividades instrumentales
        "instrumentales_nivel_apoyo": _NIVEL_APOYO,
        "instrumentales_actividades": TEXTO_CORTO,
        "instrumentales_criar_apoyo": TEXTO_CORTO,
        "instrumentales_comunicacion_apoyo": TEXTO_CORTO,
        "instrumentales_movilidad_apoyo": TEXTO_CORTO,
        "instrumentales_finanzas": TEXTO_CORTO,
        "instrumentales_cocina_limpieza": TEXTO_CORTO,
        "instrumentales_crear_hogar": TEXTO_CORTO,
        "instrumentales_salud_cuenta_apoyo": TEXTO_CORTO,
        "instrumentales_nota": TEXTO_CORTO,
        # Actividades
        "actividades_nivel_apoyo": _NIVEL_APOYO,
        "actividades_apoyo": TEXTO_CORTO,
        "actividades_esparcimiento_apoyo": TEXTO_CORTO,
        "actividades_esparcimiento_cuenta_apoyo": TEXTO_CORTO,
        "actividades_complementarios_apoyo": TEXTO_CORTO,
        "actividades_complementarios_cuenta_apoyo": TEXTO_CORTO,
        "actividades_subsidios_cuenta_apoyo": TEXTO_CORTO,
        "actividades_nota": TEXTO_CORTO,
        # Discriminacion
        "discriminacion_nivel_apoyo": _NIVEL_APOYO,
        "discriminacion": TEXTO_CORTO,
        "discriminacion_violencia_apoyo": TEXTO_CORTO,
        "discriminacion_violencia_cuenta_apoyo": TEXTO_CORTO,
        "discriminacion_vulneracion_apoyo": TEXTO_CORTO,
        "discriminacion_vulneracion_cuenta_apoyo": TEXTO_CORTO,
        "discriminacion_nota": TEXTO_CORTO,
    }


def run():
    # --- Limpiar cache ---
    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    # --- Sección 1: Datos generales ---
    s1 = dict(SECTION_1_BASE)
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # --- Sección 2: Oferentes ---
    mod.FORM_CACHE["section_2"] = [_build_oferente()]

    # --- Sección 5: Ajustes ---
    mod.FORM_CACHE["section_5"] = {
        "ajustes": TEXTO_LARGO,
        "ajustes_recomendaciones": TEXTO_LARGO,
        "nota": TEXTO_CORTO,
    }

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
