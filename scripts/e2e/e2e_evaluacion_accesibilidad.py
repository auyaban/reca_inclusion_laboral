#!/usr/bin/env python
"""
E2E — Evaluacion de Accesibilidad
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

from formularios.evaluacion_programa import evaluacion_accesibilidad as mod
from scripts.e2e.e2e_constants import SECTION_1_BASE, TEXTO_CORTO, TEXTO_LARGO, ASISTENTES

FORM_ID = "evaluacion_accesibilidad"

# Primera opción de accesible para ítems tipo "accesible_con_observaciones"
_ACC = "Si"
# Primera opción de respuesta genérica
_RESP = "Si"


def run():
    # --- Limpiar cache ---
    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    # --- Sección 1: Datos generales ---
    s1 = dict(SECTION_1_BASE)
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # --- Sección 2.1: Condiciones de movilidad y urbanísticas ---
    mod.FORM_CACHE["section_2_1"] = {
        "transporte_publico_accesible": _ACC, "transporte_publico_observaciones": TEXTO_CORTO,
        "rutas_pcd_accesible": _ACC, "rutas_pcd_observaciones": TEXTO_CORTO,
        "parqueaderos_accesible": _ACC, "parqueaderos_observaciones": TEXTO_CORTO,
        "ubicacion_accesible_accesible": _ACC, "ubicacion_accesible_observaciones": TEXTO_CORTO,
        "vias_cercanas_accesible": _ACC, "vias_cercanas_observaciones": TEXTO_CORTO,
        "paso_peatonal_accesible": _ACC, "paso_peatonal_observaciones": TEXTO_CORTO,
        "rampas_cerca_accesible": _ACC, "rampas_cerca_observaciones": TEXTO_CORTO,
        "senales_podotactiles_accesible": _ACC, "senales_podotactiles": TEXTO_CORTO,
        "alumbrado_publico_accesible": _ACC, "alumbrado_publico": TEXTO_CORTO,
    }

    # --- Sección 2.2: Condiciones de accesibilidad general ---
    mod.FORM_CACHE["section_2_2"] = {
        "areas_administrativa_operativa_accesible": _ACC, "areas_administrativa_operativa": TEXTO_CORTO,
        "zonas_comunes_accesible": _ACC, "zonas_comunes_observaciones": TEXTO_CORTO,
        "enfermeria_accesible_accesible": _ACC, "enfermeria_accesible": TEXTO_CORTO,
        "enfermeria_accesible_observaciones": TEXTO_CORTO,
        "ergonomia_administrativa_accesible": _ACC, "ergonomia_administrativa": TEXTO_CORTO,
        "ergonomia_administrativa_observaciones": TEXTO_CORTO,
        "ergonomia_operativa_accesible": _ACC, "ergonomia_operativa": TEXTO_CORTO,
        "ergonomia_operativa_observaciones": TEXTO_CORTO,
        "mobiliario_zonas_comunes_accesible": _ACC, "mobiliario_zonas_comunes": TEXTO_CORTO,
        "mobiliario_zonas_comunes_observaciones": TEXTO_CORTO,
        "evaluacion_ergonomica_puestos_accesible": _ACC, "evaluacion_ergonomica_puestos": TEXTO_CORTO,
        "ventilacion_area_administrativa_accesible": _ACC, "ventilacion_area_administrativa": TEXTO_CORTO,
        "ventilacion_area_administrativa_secundaria": TEXTO_CORTO,
        "ventilacion_area_operativa_accesible": _ACC, "ventilacion_area_operativa": TEXTO_CORTO,
        "ventilacion_area_operativa_secundaria": TEXTO_CORTO,
        "ventilacion_areas_comunes_accesible": _ACC, "ventilacion_areas_comunes": TEXTO_CORTO,
        "ventilacion_areas_comunes_secundaria": TEXTO_CORTO,
        "iluminacion_area_administrativa_accesible": _ACC, "iluminacion_area_administrativa": TEXTO_CORTO,
        "iluminacion_area_administrativa_secundaria": TEXTO_CORTO,
        "iluminacion_area_operativa_accesible": _ACC, "iluminacion_area_operativa": TEXTO_CORTO,
        "iluminacion_area_operativa_secundaria": TEXTO_CORTO,
        "iluminacion_areas_comunes_accesible": _ACC, "iluminacion_areas_comunes": TEXTO_CORTO,
        "iluminacion_areas_comunes_secundaria": TEXTO_CORTO,
        "ruido_area_administrativa_accesible": _ACC, "ruido_area_administrativa": TEXTO_CORTO,
        "ruido_area_administrativa_secundaria": TEXTO_CORTO, "ruido_area_administrativa_terciaria": TEXTO_CORTO,
        "ruido_area_operativa_accesible": _ACC, "ruido_area_operativa": TEXTO_CORTO,
        "ruido_area_operativa_secundaria": TEXTO_CORTO, "ruido_area_operativa_terciaria": TEXTO_CORTO,
        "ruido_areas_comunes_accesible": _ACC, "ruido_areas_comunes": TEXTO_CORTO,
        "ruido_areas_comunes_secundaria": TEXTO_CORTO, "ruido_areas_comunes_terciaria": TEXTO_CORTO,
        "flexibilidad_hibrido_remoto_accesible": _ACC, "flexibilidad_hibrido_remoto": TEXTO_CORTO,
        "flexibilidad_horarios_calamidades_accesible": _ACC, "flexibilidad_horarios_calamidades": TEXTO_CORTO,
        "sala_lactancia_accesible": _ACC, "sala_lactancia": TEXTO_CORTO,
        "protocolo_sala_lactancia_accesible": _ACC, "protocolo_sala_lactancia": TEXTO_CORTO,
        "linea_purpura_accesible": _ACC, "linea_purpura": TEXTO_CORTO,
        "salas_amigas_accesible": _ACC, "salas_amigas": TEXTO_CORTO,
        "protocolo_hostigamiento_acoso_sexual_accesible": _ACC, "protocolo_hostigamiento_acoso_sexual": TEXTO_CORTO,
        "protocolo_acoso_laboral_accesible": _ACC, "protocolo_acoso_laboral": TEXTO_CORTO,
        "practicas_equidad_genero_accesible": _ACC, "practicas_equidad_genero": TEXTO_CORTO,
        "canales_comunicacion_lenguaje_inclusivo_accesible": _ACC, "canales_comunicacion_lenguaje_inclusivo": TEXTO_CORTO,
    }

    # --- Sección 2.3: Condiciones discapacidad física ---
    mod.FORM_CACHE["section_2_3"] = {
        "entrada_salida_accesible": _ACC, "entrada_salida_observaciones": TEXTO_CORTO,
        "rampas_interior_usr_accesible": _ACC, "rampas_interior_usr_observaciones": TEXTO_CORTO,
        "ascensor_interior_accesible": _ACC, "ascensor_interior_observaciones": TEXTO_CORTO,
        "zonas_oficinas_accesibles_accesible": _ACC, "zonas_oficinas_accesibles_observaciones": TEXTO_CORTO,
        "cafeteria_accesible_accesible": _ACC, "cafeteria_accesible_observaciones": TEXTO_CORTO,
        "zonas_descanso_accesibles_accesible": _ACC, "zonas_descanso_accesibles_observaciones": TEXTO_CORTO,
        "pasillos_amplios_accesible": _ACC, "pasillos_amplios_observaciones": TEXTO_CORTO,
        "escaleras_doble_funcion_accesible": _ACC, "escaleras_doble_funcion": TEXTO_CORTO,
        "escaleras_doble_funcion_secundaria": TEXTO_CORTO, "escaleras_doble_funcion_terciaria": TEXTO_CORTO,
        "escaleras_doble_funcion_cuaternaria": TEXTO_CORTO,
        "escaleras_interior_accesible": _ACC, "escaleras_interior": TEXTO_CORTO,
        "escaleras_interior_secundaria": TEXTO_CORTO, "escaleras_interior_terciaria": TEXTO_CORTO,
        "escaleras_interior_cuaternaria": TEXTO_CORTO,
        "escaleras_emergencia_accesible": _ACC, "escaleras_emergencia": TEXTO_CORTO,
        "escaleras_emergencia_secundaria": TEXTO_CORTO, "escaleras_emergencia_terciaria": TEXTO_CORTO,
        "bano_discapacidad_fisica_accesible": _ACC, "bano_discapacidad_fisica": TEXTO_CORTO,
        "bano_discapacidad_fisica_secundaria": TEXTO_CORTO, "bano_discapacidad_fisica_terciaria": TEXTO_CORTO,
        "bano_discapacidad_fisica_cuaternaria": TEXTO_CORTO, "bano_discapacidad_fisica_quinary": TEXTO_CORTO,
        "silla_evacuacion_usr_accesible": _ACC, "silla_evacuacion_usr": TEXTO_CORTO,
        "silla_evacuacion_oruga_accesible": _ACC, "silla_evacuacion_oruga": TEXTO_CORTO,
        "ergonomia_superficies_irregulares_accesible": _ACC, "ergonomia_superficies_irregulares": TEXTO_CORTO,
        "senalizacion_ntc_accesible": _ACC, "senalizacion_ntc": TEXTO_CORTO,
        "senalizacion_ntc_secundaria": TEXTO_CORTO,
        "mapa_evacuacion_ntc_accesible": _ACC, "mapa_evacuacion_ntc": TEXTO_CORTO,
        "mapa_evacuacion_ntc_secundaria": TEXTO_CORTO, "mapa_evacuacion_ntc_terciaria": TEXTO_CORTO,
        "ajustes_razonables_individualizados_accesible": _ACC, "ajustes_razonables_individualizados": TEXTO_CORTO,
        "ajustes_razonables_detalle": TEXTO_CORTO,
    }

    # --- Sección 2.4: Condiciones discapacidad sensorial ---
    mod.FORM_CACHE["section_2_4"] = {
        "senalizacion_orientacion_accesible": _ACC, "senalizacion_orientacion": TEXTO_CORTO,
        "senalizacion_emergencia_accesible": _ACC, "senalizacion_emergencia": TEXTO_CORTO,
        "distribucion_zonas_comunes_accesible": _ACC, "distribucion_zonas_comunes_observaciones": TEXTO_CORTO,
        "senalizacion_mapa_evacuacion_accesible": _ACC, "senalizacion_mapa_evacuacion_observaciones": TEXTO_CORTO,
        "ascensor_apoyo_visual_sonoro_accesible": _ACC, "ascensor_apoyo_visual_sonoro_observaciones": TEXTO_CORTO,
        "apoyo_seguridad_ubicacion_accesible": _ACC, "apoyo_seguridad_ubicacion_observaciones": TEXTO_CORTO,
        "senalizacion_ntc_accesible": _ACC, "senalizacion_ntc": TEXTO_CORTO,
        "senalizacion_ntc_secundaria": TEXTO_CORTO,
        "mapa_evacuacion_ntc_accesible": _ACC, "mapa_evacuacion_ntc": TEXTO_CORTO,
        "mapa_evacuacion_ntc_secundaria": TEXTO_CORTO, "mapa_evacuacion_ntc_terciaria": TEXTO_CORTO,
        "informacion_accesible_ingreso_accesible": _ACC, "informacion_accesible_ingreso": TEXTO_CORTO,
        "informacion_accesible_ingreso_secundaria": TEXTO_CORTO,
        "informacion_accesible_ingreso_terciaria": TEXTO_CORTO,
        "informacion_accesible_ingreso_cuaternaria": TEXTO_CORTO,
        "medios_tecnologicos_ingreso_accesible": _ACC, "medios_tecnologicos_ingreso_observaciones": TEXTO_CORTO,
        "material_seleccion_accesible_accesible": _ACC, "material_seleccion_accesible": TEXTO_CORTO,
        "material_seleccion_accesible_secundaria": TEXTO_CORTO, "material_seleccion_accesible_terciaria": TEXTO_CORTO,
        "material_contratacion_accesible_accesible": _ACC, "material_contratacion_accesible": TEXTO_CORTO,
        "material_contratacion_accesible_secundaria": TEXTO_CORTO,
        "material_induccion_accesible_accesible": _ACC, "material_induccion_accesible": TEXTO_CORTO,
        "material_induccion_accesible_secundaria": TEXTO_CORTO, "material_induccion_accesible_terciaria": TEXTO_CORTO,
        "material_evaluacion_desempeno_accesible": _ACC, "material_evaluacion_desempeno": TEXTO_CORTO,
        "material_evaluacion_desempeno_secundaria": TEXTO_CORTO, "material_evaluacion_desempeno_terciaria": TEXTO_CORTO,
        "plataformas_autogestion_accesible": _ACC, "plataformas_autogestion": TEXTO_CORTO,
        "plataformas_autogestion_secundaria": TEXTO_CORTO, "plataformas_autogestion_terciaria": TEXTO_CORTO,
        "plataformas_autogestion_cuaternaria": TEXTO_CORTO,
        "alarma_emergencia_accesible": _ACC, "alarma_emergencia": TEXTO_CORTO,
        "ajustes_razonables_individualizados_accesible": _ACC, "ajustes_razonables_individualizados": TEXTO_CORTO,
        "ajustes_razonables_individualizados_detalle": TEXTO_CORTO,
    }

    # --- Sección 2.5: Condiciones discapacidad intelectual / TEA ---
    mod.FORM_CACHE["section_2_5"] = {
        "material_seleccion_cognitiva_accesible": _ACC, "material_seleccion_cognitiva": TEXTO_CORTO,
        "material_contratacion_cognitiva_accesible": _ACC, "material_contratacion_cognitiva": TEXTO_CORTO,
        "material_induccion_cognitiva_accesible": _ACC, "material_induccion_cognitiva": TEXTO_CORTO,
        "material_evaluacion_cognitiva_accesible": _ACC, "material_evaluacion_cognitiva": TEXTO_CORTO,
        "ascensor_facil_ubicacion_accesible": _ACC, "ascensor_facil_ubicacion_observaciones": TEXTO_CORTO,
        "distribucion_zonas_comunes_percepcion_accesible": _ACC,
        "distribucion_zonas_comunes_percepcion_observaciones": TEXTO_CORTO,
        "plataformas_autogestion_intelectual_accesible": _ACC, "plataformas_autogestion_intelectual": TEXTO_CORTO,
        "plataformas_autogestion_intelectual_secundaria": TEXTO_CORTO,
        "ajustes_razonables_intelectual_accesible": _ACC, "ajustes_razonables_intelectual": TEXTO_CORTO,
        "ajustes_razonables_intelectual_detalle": TEXTO_CORTO,
    }

    # --- Sección 2.6: Condiciones discapacidad psicosocial ---
    mod.FORM_CACHE["section_2_6"] = {
        "ajustes_razonables_psicosocial_accesible": _ACC,
        "ajustes_razonables_psicosocial": TEXTO_CORTO,
        "ajustes_razonables_psicosocial_detalle": TEXTO_CORTO,
    }

    # --- Sección 3: Condiciones organizacionales (flat dict) ---
    mod.FORM_CACHE["section_3"] = {
        "experiencia_vinculacion_pcd_accesible": _ACC, "experiencia_vinculacion_pcd_observaciones": TEXTO_CORTO,
        "personal_tercerizado_capacitado_accesible": _ACC, "personal_tercerizado_capacitado": TEXTO_CORTO,
        "personal_directo_capacitado_accesible": _ACC, "personal_directo_capacitado": TEXTO_CORTO,
        "apoyo_arl_seguridad_accesible": _ACC, "apoyo_arl_seguridad": TEXTO_CORTO,
        "capacitacion_emergencias_accesible": _ACC, "capacitacion_emergencias_observaciones": TEXTO_CORTO,
        "politica_diversidad_inclusion_accesible": _ACC, "politica_diversidad_inclusion": TEXTO_CORTO,
        "rrhh_normatividad_accesible": _ACC, "rrhh_normatividad": TEXTO_CORTO,
        "rrhh_normatividad_secundaria": TEXTO_CORTO, "rrhh_normatividad_terciaria": TEXTO_CORTO,
        "rrhh_normatividad_cuaternaria": TEXTO_CORTO,
        "ajustes_razonables_empresa_accesible": _ACC, "ajustes_razonables_empresa": TEXTO_CORTO,
        "ajustes_razonables_empresa_secundaria": TEXTO_CORTO, "ajustes_razonables_empresa_terciaria": TEXTO_CORTO,
        "ajustes_razonables_empresa_cuaternaria": TEXTO_CORTO,
        "protocolo_emergencias_pcd_accesible": _ACC, "protocolo_emergencias_pcd": TEXTO_CORTO,
        "protocolo_emergencias_pcd_secundaria": TEXTO_CORTO, "protocolo_emergencias_pcd_terciaria": TEXTO_CORTO,
        "apoyo_bomberos_discapacidad_accesible": _ACC, "apoyo_bomberos_discapacidad": TEXTO_CORTO,
        "apoyo_bomberos_discapacidad_detalle": TEXTO_CORTO,
        "disponibilidad_tiempo_inclusion_accesible": _ACC, "disponibilidad_tiempo_inclusion": TEXTO_CORTO,
        "practicas_equidad_genero_accesible": _ACC, "practicas_equidad_genero": TEXTO_CORTO,
    }

    # --- Sección 4: Concepto de la evaluación ---
    mod.FORM_CACHE["section_4"] = {"nivel_accesibilidad": "Alto"}

    # --- Sección 5: Ajustes razonables (flat dict por tipo de discapacidad) ---
    mod.FORM_CACHE["section_5"] = {
        "discapacidad_fisica_ajustes": TEXTO_LARGO, "discapacidad_fisica": TEXTO_CORTO,
        "discapacidad_fisica_nota": TEXTO_CORTO,
        "discapacidad_fisica_usr_ajustes": TEXTO_LARGO, "discapacidad_fisica_usr": TEXTO_CORTO,
        "discapacidad_fisica_usr_nota": TEXTO_CORTO,
        "discapacidad_auditiva_ajustes": TEXTO_LARGO, "discapacidad_auditiva": TEXTO_CORTO,
        "discapacidad_auditiva_nota": TEXTO_CORTO,
        "discapacidad_visual_ajustes": TEXTO_LARGO, "discapacidad_visual": TEXTO_CORTO,
        "discapacidad_visual_nota": TEXTO_CORTO,
        "discapacidad_intelectual_ajustes": TEXTO_LARGO, "discapacidad_intelectual": TEXTO_CORTO,
        "discapacidad_intelectual_nota": TEXTO_CORTO,
        "trastorno_espectro_autista_ajustes": TEXTO_LARGO, "trastorno_espectro_autista": TEXTO_CORTO,
        "trastorno_espectro_autista_nota": TEXTO_CORTO,
        "discapacidad_psicosocial_ajustes": TEXTO_LARGO, "discapacidad_psicosocial": TEXTO_CORTO,
        "discapacidad_psicosocial_nota": TEXTO_CORTO,
        "discapacidad_visual_baja_vision_ajustes": TEXTO_LARGO, "discapacidad_visual_baja_vision": TEXTO_CORTO,
        "discapacidad_visual_baja_vision_nota": TEXTO_CORTO,
        "discapacidad_auditiva_reducida_ajustes": TEXTO_LARGO, "discapacidad_auditiva_reducida": TEXTO_CORTO,
        "discapacidad_auditiva_reducida_nota": TEXTO_CORTO,
    }

    # --- Sección 6: Observaciones generales ---
    mod.FORM_CACHE["section_6"] = {"observaciones_generales": TEXTO_LARGO}

    # --- Sección 7: Cargos compatibles ---
    mod.FORM_CACHE["section_7"] = {"cargos_compatibles": TEXTO_CORTO}

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
