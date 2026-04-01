"""
End-to-end test: fill each form module with minimal test data,
export to Google Sheets, then read back key cells to verify mapping.
"""
import sys
import os
import io
import json
import time
import traceback

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

from google_sheets_client import read_sheet_values, extract_spreadsheet_id

RESULTS = []
CREATED_SHEETS = []


def read_cell(spreadsheet_id, sheet_name, cell):
    """Read a single cell value from a sheet."""
    range_str = f"'{sheet_name}'!{cell}"
    vals = read_sheet_values(spreadsheet_id, range_str)
    if vals and vals[0]:
        return vals[0][0]
    return None


def verify_cells(test_name, spreadsheet_id, sheet_name, expected_map):
    """Verify a dict of {cell: expected_value} against the sheet."""
    errors = []
    ok_count = 0
    for cell, expected in expected_map.items():
        actual = read_cell(spreadsheet_id, sheet_name, cell)
        actual_str = str(actual or "").strip()
        expected_str = str(expected or "").strip()
        if actual_str != expected_str:
            errors.append(f"  {cell}: expected={expected_str!r} got={actual_str!r}")
        else:
            ok_count += 1
    status = "PASS" if not errors else "FAIL"
    RESULTS.append((test_name, status, ok_count, len(expected_map), errors))
    print(f"  [{status}] {test_name}: {ok_count}/{len(expected_map)} cells OK")
    for e in errors:
        print(e)
    return len(errors) == 0


def extract_id_from_result(result):
    """Extract spreadsheet ID from export result."""
    url = result.get("output_path", "")
    if "docs.google.com" in url:
        return extract_spreadsheet_id(url)
    return result.get("drive_file_id", "")


# ═══════════════════════════════════════════════════════════════════
# TEST 1: SENSIBILIZACION
# ═══════════════════════════════════════════════════════════════════
def test_sensibilizacion():
    print("\n=== TEST: sensibilizacion ===")
    from formularios.sensibilizacion import sensibilizacion as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Virtual",
        "nombre_empresa": "Empresa Test E2E",
        "ciudad_empresa": "Bogota",
        "direccion_empresa": "Calle 123",
        "nit_empresa": "900111222",
        "correo_1": "test@empresa.com",
        "telefono_empresa": "3001234567",
        "contacto_empresa": "Juan Perez",
        "cargo": "Gerente RRHH",
        "caja_compensacion": "Compensar",
        "sede_empresa": "Principal",
        "asesor": "Maria Asesora",
        "profesional_asignado": "Pedro Profesional",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    s3 = {"observaciones": "Observacion de prueba E2E"}
    mod.FORM_CACHE["section_3"] = s3

    s5 = [
        {"nombre": "Asistente Uno", "cargo": "Coordinador"},
        {"nombre": "Asistente Dos", "cargo": "Analista"},
    ]
    mod.FORM_CACHE["section_5"] = s5

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("sensibilizacion", sid, sheet, {
        "D7": "26/03/2026",
        "N7": "Virtual",
        "D8": "Empresa Test E2E",
        "N8": "Bogota",
        "D9": "Calle 123",
        "N9": "900111222",
        "D10": "test@empresa.com",
        "N10": "3001234567",
        "D11": "Juan Perez",
        "N11": "Gerente RRHH",
        "D12": "Maria Asesora",
        "N12": "Principal",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 2: PRESENTACION PROGRAMA
# ═══════════════════════════════════════════════════════════════════
def test_presentacion_programa():
    print("\n=== TEST: presentacion_programa ===")
    from formularios.presentacion_programa import presentacion_programa as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Presencial",
        "nombre_empresa": "Empresa PresentProg",
        "ciudad_empresa": "Medellin",
        "direccion_empresa": "Carrera 45",
        "nit_empresa": "800222333",
        "correo_1": "pp@test.com",
        "telefono_empresa": "3109876543",
        "contacto_empresa": "Ana Lopez",
        "cargo": "Directora",
        "caja_compensacion": "Colsubsidio",
        "sede_empresa": "Norte",
        "asesor": "Carlos Asesor",
        "profesional_asignado": "Laura Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    s3 = {"observaciones": "Obs presentacion test"}
    mod.FORM_CACHE["section_3"] = s3

    s5 = [{"nombre": "Persona Test", "cargo": "Cargo Test"}]
    mod.FORM_CACHE["section_5"] = s5

    result = mod.export_to_excel()
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAMES["presentacion"]

    verify_cells("presentacion_programa", sid, sheet, {
        "D7": "26/03/2026",
        "Q7": "Presencial",
        "D8": "Empresa PresentProg",
        "Q8": "Medellin",
        "D13": "Laura Prof",
        "D14": "Carlos Asesor",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 3: INDUCCION ORGANIZACIONAL
# ═══════════════════════════════════════════════════════════════════
def test_induccion_organizacional():
    print("\n=== TEST: induccion_organizacional ===")
    from formularios.induccion_organizacional import induccion_organizacional as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Mixta",
        "nombre_empresa": "Empresa IndOrg",
        "ciudad_empresa": "Cali",
        "direccion_empresa": "Av 10",
        "nit_empresa": "700333444",
        "correo_1": "io@test.com",
        "telefono_empresa": "3201112233",
        "contacto_empresa": "Diego Ramirez",
        "cargo": "Jefe RRHH",
        "caja_compensacion": "Cafam",
        "sede_empresa": "Sur",
        "asesor": "Sofia Asesora",
        "profesional_asignado": "Roberto Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1
    mod.FORM_CACHE["section_3"] = {"observaciones": "Obs induccion org"}
    mod.FORM_CACHE["section_5"] = {"observaciones": "Obs section 5 IO"}
    mod.FORM_CACHE["section_6"] = [{"nombre": "Test IO", "cargo": "Test Cargo"}]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("induccion_organizacional", sid, sheet, {
        "D7": "2026-03-26",
        "N7": "Mixta",
        "D8": "Empresa IndOrg",
        "N9": "700333444",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 4: INDUCCION OPERATIVA
# ═══════════════════════════════════════════════════════════════════
def test_induccion_operativa():
    print("\n=== TEST: induccion_operativa ===")
    from formularios.induccion_operativa import induccion_operativa as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Virtual",
        "nombre_empresa": "Empresa IndOp",
        "ciudad_empresa": "Barranquilla",
        "direccion_empresa": "Calle 80",
        "nit_empresa": "600444555",
        "correo_1": "iop@test.com",
        "telefono_empresa": "3005556677",
        "contacto_empresa": "Lucia Gomez",
        "cargo": "Coord RRHH",
        "caja_compensacion": "Comfenalco",
        "sede_empresa": "Centro",
        "asesor": "Pablo Asesor",
        "profesional_asignado": "Diana Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1
    mod.FORM_CACHE["section_3"] = {"observaciones": "Obs induccion op"}
    mod.FORM_CACHE["section_5"] = [{"nombre": "Test IOP", "cargo": "Cargo IOP"}]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("induccion_operativa", sid, sheet, {
        "E7": "26/3/2026",
        "M7": "Virtual",
        "E8": "Empresa IndOp",
        "M8": "Barranquilla",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 5: EVALUACION ACCESIBILIDAD
# ═══════════════════════════════════════════════════════════════════
def test_evaluacion_accesibilidad():
    print("\n=== TEST: evaluacion_accesibilidad ===")
    from formularios.evaluacion_programa import evaluacion_accesibilidad as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Presencial",
        "nombre_empresa": "Empresa EvalAcces",
        "ciudad_empresa": "Bucaramanga",
        "direccion_empresa": "Cra 27",
        "nit_empresa": "500555666",
        "correo_1": "ea@test.com",
        "telefono_empresa": "3116667788",
        "contacto_empresa": "Fernando Torres",
        "cargo": "Director SST",
        "caja_compensacion": "Compensar",
        "sede_empresa": "Unica",
        "asesor": "Gloria Asesora",
        "profesional_asignado": "Ricardo Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # Minimal section 2 data
    s2 = {
        "nombre_profesional": "Ricardo Prof",
        "actividad_economica": "Tecnologia",
        "numero_trabajadores": "50",
    }
    mod.FORM_CACHE["section_2"] = s2

    result = mod.export_to_excel()
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("evaluacion_accesibilidad", sid, sheet, {
        "D7": "26/3/2026",
        "P7": "Presencial",
        "D8": "Empresa EvalAcces",
        "P8": "Bucaramanga",
        "P9": "500555666",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 6: CONDICIONES VACANTE
# ═══════════════════════════════════════════════════════════════════
def test_condiciones_vacante():
    print("\n=== TEST: condiciones_vacante ===")
    from formularios.condiciones_vacante import condiciones_vacante as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Virtual",
        "nombre_empresa": "Empresa CondVac",
        "ciudad_empresa": "Pereira",
        "direccion_empresa": "Calle 19",
        "nit_empresa": "400666777",
        "correo_1": "cv@test.com",
        "telefono_empresa": "3007778899",
        "contacto_empresa": "Sandra Ruiz",
        "cargo": "Coord Seleccion",
        "caja_compensacion": "Comfandi",
        "sede_empresa": "Central",
        "asesor": "Andres Asesor",
        "profesional_asignado": "Monica Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # Minimal other sections
    mod.FORM_CACHE["section_2"] = {
        "nombre_cargo": "Operario",
        "area_trabajo": "Produccion",
    }
    mod.FORM_CACHE["section_3"] = {}
    mod.FORM_CACHE["section_4"] = {}
    mod.FORM_CACHE["section_5"] = {}
    mod.FORM_CACHE["section_7"] = {"observaciones_recomendaciones": "Recomendacion de prueba E2E"}
    mod.FORM_CACHE["section_8"] = [
        {"nombre": "Asistente CV1", "cargo": "Gerente"},
    ]

    result = mod.export_to_excel()
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("condiciones_vacante", sid, sheet, {
        "F7": "26/03/2026",
        "N7": "Virtual",
        "F8": "Empresa CondVac",
        "N8": "Pereira",
        "F9": "Calle 19",
        "N9": "400666777",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 7: SELECCION INCLUYENTE LABS
# ═══════════════════════════════════════════════════════════════════
def test_seleccion_incluyente_labs():
    print("\n=== TEST: seleccion_incluyente_labs ===")
    from formularios.seleccion_incluyente_labs import seleccion_incluyente as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Presencial",
        "nombre_empresa": "Empresa SelIncLabs",
        "ciudad_empresa": "Manizales",
        "direccion_empresa": "Av Santander",
        "nit_empresa": "300777888",
        "correo_1": "sil@test.com",
        "telefono_empresa": "3128889900",
        "contacto_empresa": "Mario Diaz",
        "cargo": "Jefe Talento",
        "caja_compensacion": "Comfamiliar",
        "sede_empresa": "Oeste",
        "asesor": "Patricia Asesora",
        "profesional_asignado": "Jorge Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    mod.FORM_CACHE["section_2"] = [{
        "numero": "1",
        "nombre_oferente": "Oferente Labs Test",
        "cedula": "1234567890",
        "discapacidad": "Discapacidad física",
    }]

    mod.FORM_CACHE["section_5"] = {
        "ajustes_recomendaciones": "Ajustes Labs E2E",
        "nota": "Nota Labs E2E",
    }

    mod.FORM_CACHE["section_6"] = [
        {"nombre": "Asistente Labs", "cargo": "Coord Labs"},
    ]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("seleccion_incluyente_labs", sid, sheet, {
        "F7": "2026-03-26",
        "N7": "Presencial",
        "F8": "Empresa SelIncLabs",
        "N8": "Manizales",
        "N9": "300777888",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 8: SELECCION INCLUYENTE - INDIVIDUAL
# ═══════════════════════════════════════════════════════════════════
def test_seleccion_incluyente_individual():
    print("\n=== TEST: seleccion_incluyente INDIVIDUAL ===")
    from formularios.seleccion_incluyente import seleccion_incluyente as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Virtual",
        "nombre_empresa": "Empresa SelInc Indiv",
        "ciudad_empresa": "Bogota",
        "direccion_empresa": "Calle 100",
        "nit_empresa": "200888999",
        "correo_1": "si@test.com",
        "telefono_empresa": "3009990011",
        "contacto_empresa": "Clara Mendez",
        "cargo": "Dir Talento",
        "caja_compensacion": "Compensar",
        "sede_empresa": "Principal",
        "asesor": "Felipe Asesor",
        "profesional_asignado": "Natalia Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # Individual: 1 oferente
    mod.FORM_CACHE["section_2"] = [{
        "numero": "1",
        "nombre_oferente": "Oferente Individual Test",
        "cedula": "9876543210",
        "certificado_porcentaje": "45",
        "discapacidad": "Discapacidad auditiva",
        "telefono_oferente": "3111234567",
        "cargo_oferente": "Asistente Admin",
        "medicamentos_nivel_apoyo": "0. No requiere apoyo.",
    }]

    mod.FORM_CACHE["section_5"] = {
        "ajustes_recomendaciones": "Ajustes seleccion individual E2E",
        "nota": "Nota seleccion individual E2E",
    }

    mod.FORM_CACHE["section_6"] = [
        {"nombre": "Asistente SI1", "cargo": "Coord"},
        {"nombre": "Asistente SI2", "cargo": "Analista"},
    ]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("seleccion_incluyente_individual", sid, sheet, {
        "F7": "2026-03-26",
        "N7": "Virtual",
        "F8": "Empresa SelInc Indiv",
        "N8": "Bogota",
        "N9": "200888999",
        # Section 2 individual cell map
        "A17": "1",  # numero
        "C17": "Oferente Individual Test",  # nombre_oferente
        "H17": "9876543210",  # cedula
        # Section 5 (individual: ajustes=79, nota=80)
        "A79": "Ajustes seleccion individual E2E",
        "A80": "Nota: Nota seleccion individual E2E",
        # Section 6 (individual start=85)
        "E85": "Asistente SI1",
        "M85": "Coord",
        "E86": "Asistente SI2",
        "M86": "Analista",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 9: SELECCION INCLUYENTE - GRUPAL (2 oferentes)
# ═══════════════════════════════════════════════════════════════════
def test_seleccion_incluyente_grupal():
    print("\n=== TEST: seleccion_incluyente GRUPAL ===")
    from formularios.seleccion_incluyente import seleccion_incluyente as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Presencial",
        "nombre_empresa": "Empresa SelInc Grupal",
        "ciudad_empresa": "Medellin",
        "direccion_empresa": "Carrera 50",
        "nit_empresa": "100999000",
        "correo_1": "sig@test.com",
        "telefono_empresa": "3040001122",
        "contacto_empresa": "Sergio Vargas",
        "cargo": "Gerente",
        "asesor": "Rita Asesora",
        "sede_empresa": "Este",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    # Group: 2 oferentes
    mod.FORM_CACHE["section_2"] = [
        {
            "numero": "1",
            "nombre_oferente": "Oferente Grupal 1",
            "cedula": "1111111111",
            "discapacidad": "Discapacidad visual baja visión",
            "desarrollo_actividad": "Desarrollo actividad compartida E2E",
            "medicamentos_nivel_apoyo": "1. Nivel de apoyo Bajo.",
        },
        {
            "numero": "2",
            "nombre_oferente": "Oferente Grupal 2",
            "cedula": "2222222222",
            "discapacidad": "Discapacidad intelectual",
            "medicamentos_nivel_apoyo": "0. No requiere apoyo.",
        },
    ]

    mod.FORM_CACHE["section_5"] = {
        "ajustes_recomendaciones": "Ajustes grupal E2E",
        "nota": "Nota grupal E2E",
    }

    mod.FORM_CACHE["section_6"] = [
        {"nombre": "Asistente SIG", "cargo": "Jefe"},
    ]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    # Group cell map: first block starts at row 16, second at 16+61=77
    # Section 1 group mapping has asesor at F12 (not F13 like individual)
    verify_cells("seleccion_incluyente_grupal", sid, sheet, {
        "F7": "2026-03-26",
        "N7": "Presencial",
        "F8": "Empresa SelInc Grupal",
        "N8": "Medellin",
        # OFERENTE 1 title
        "A16": "OFERENTE 1",
        # OFERENTE 1 data (group first block cell map: nombre at C19)
        "C19": "Oferente Grupal 1",
        "H19": "1111111111",
        # Shared actividad
        "A14": "Desarrollo actividad compartida E2E",
        # OFERENTE 2 title (row 16 + 61 = 77)
        "A77": "OFERENTE 2",
        # Section 5 (group base: ajustes=139, nota=140)
        "A139": "Ajustes grupal E2E",
        "A140": "Nota: Nota grupal E2E",
        # Section 6 (group base start=145)
        "E145": "Asistente SIG",
        "M145": "Jefe",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 10: CONTRATACION INCLUYENTE - INDIVIDUAL
# ═══════════════════════════════════════════════════════════════════
def test_contratacion_individual():
    print("\n=== TEST: contratacion_incluyente INDIVIDUAL ===")
    from formularios.contratacion_incluyente import contratacion_incluyente as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Virtual",
        "nombre_empresa": "Empresa ContInc Indiv",
        "ciudad_empresa": "Bogota",
        "direccion_empresa": "Calle 26",
        "nit_empresa": "900123456",
        "correo_1": "ci@test.com",
        "telefono_empresa": "3051112233",
        "contacto_empresa": "Rosa Jimenez",
        "cargo": "Dir RRHH",
        "caja_compensacion": "Compensar",
        "sede_empresa": "Centro",
        "asesor": "Oscar Asesor",
        "profesional_asignado": "Carmen Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    mod.FORM_CACHE["section_2"] = [{
        "numero": "1",
        "nombre_oferente": "Vinculado Individual Test",
        "cedula": "5555555555",
        "certificado_porcentaje": "60",
        "discapacidad": "Discapacidad física",
        "telefono_oferente": "3161234567",
        "cargo_oferente": "Auxiliar",
        "tipo_contrato": "Contrato de trabajo a término fijo",
        "desarrollo_actividad": "Actividad individual test",
        "contrato_lee_nivel_apoyo": "0. No requiere apoyo.",
    }]

    mod.FORM_CACHE["section_6"] = {
        "ajustes_recomendaciones": "Ajustes contratacion individual E2E",
    }

    mod.FORM_CACHE["section_7"] = [
        {"nombre": "Asistente CI", "cargo": "Coordinador"},
    ]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    verify_cells("contratacion_individual", sid, sheet, {
        "D7": "2026-03-26",
        "L7": "Virtual",
        "D8": "Empresa ContInc Indiv",
        "L8": "Bogota",
        "L9": "900123456",
        # Section 2 individual: nombre at C18, cedula at H18
        "C18": "Vinculado Individual Test",
        "H18": "5555555555",
        # Section 6 (individual ajustes=69)
        "A69": "Ajustes contratacion individual E2E",
        # Section 7 (individual start=74)
        "C74": "Asistente CI",
        "K74": "Coordinador",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 11: CONTRATACION INCLUYENTE - GRUPAL 2 VINCULADOS
# ═══════════════════════════════════════════════════════════════════
def test_contratacion_grupal_2():
    print("\n=== TEST: contratacion_incluyente GRUPAL 2 vinculados ===")
    from formularios.contratacion_incluyente import contratacion_incluyente as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Presencial",
        "nombre_empresa": "Empresa ContInc Grup2",
        "ciudad_empresa": "Cali",
        "direccion_empresa": "Av 6ta",
        "nit_empresa": "800234567",
        "correo_1": "cg2@test.com",
        "telefono_empresa": "3062223344",
        "contacto_empresa": "Luis Herrera",
        "cargo": "Coord Contratacion",
        "caja_compensacion": "Comfandi",
        "sede_empresa": "Zona Franca",
        "asesor": "Elena Asesora",
        "profesional_asignado": "Miguel Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    mod.FORM_CACHE["section_2"] = [
        {
            "numero": "1",
            "nombre_oferente": "Vinculado Grupal 1",
            "cedula": "6666666666",
            "discapacidad": "Discapacidad auditiva",
            "tipo_contrato": "Contrato temporal",
            "desarrollo_actividad": "Actividad grupal compartida E2E",
        },
        {
            "numero": "2",
            "nombre_oferente": "Vinculado Grupal 2",
            "cedula": "7777777777",
            "discapacidad": "Discapacidad visual baja visión",
            "tipo_contrato": "Contrato de aprendizaje",
        },
    ]

    mod.FORM_CACHE["section_6"] = {
        "ajustes_recomendaciones": "Ajustes contratacion grupal 2 E2E",
    }

    mod.FORM_CACHE["section_7"] = [
        {"nombre": "Asistente CG2-1", "cargo": "Jefe"},
        {"nombre": "Asistente CG2-2", "cargo": "Analista"},
    ]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    # Group cell map: vinculado 1 block start=19, block height=52
    # SECTION_2_GROUP_FIRST_BLOCK_CELL_MAP row adjustments: <=24 gets +5, >=30 gets +3
    # nombre_oferente original (C,18), +5 = (C,23)
    # cedula original (H,18), +5 = (H,23)
    verify_cells("contratacion_grupal_2", sid, sheet, {
        "D7": "2026-03-26",
        "L7": "Presencial",
        "D8": "Empresa ContInc Grup2",
        "L8": "Cali",
        "L9": "800234567",
        # Shared actividad
        "A15": "Actividad grupal compartida E2E",
        # Vinculado 1 data (group first block)
        "C23": "Vinculado Grupal 1",
        "H23": "6666666666",
        # Vinculado 2 data (offset by 52)
        "C75": "Vinculado Grupal 2",
        "H75": "7777777777",
        # Section 6 (group base ajustes=125)
        "A125": "Ajustes contratacion grupal 2 E2E",
        # Section 7 (group base start=131)
        "C131": "Asistente CG2-1",
        "K131": "Jefe",
        "C132": "Asistente CG2-2",
        "K132": "Analista",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# TEST 12: CONTRATACION INCLUYENTE - GRUPAL 5 VINCULADOS
# ═══════════════════════════════════════════════════════════════════
def test_contratacion_grupal_5():
    print("\n=== TEST: contratacion_incluyente GRUPAL 5 vinculados ===")
    from formularios.contratacion_incluyente import contratacion_incluyente as mod

    mod.FORM_CACHE.clear()
    mod.SECTION_1_CACHE.clear()

    s1 = {
        "fecha_visita": "2026-03-26",
        "modalidad": "Mixta",
        "nombre_empresa": "Empresa ContInc Grup5",
        "ciudad_empresa": "Cartagena",
        "direccion_empresa": "Bocagrande",
        "nit_empresa": "700345678",
        "correo_1": "cg5@test.com",
        "telefono_empresa": "3073334455",
        "contacto_empresa": "Camila Santos",
        "cargo": "Dir Gestion Humana",
        "caja_compensacion": "Comfenalco",
        "sede_empresa": "Planta",
        "asesor": "Ivan Asesor",
        "profesional_asignado": "Teresa Prof",
    }
    mod.SECTION_1_CACHE.update(s1)
    mod.FORM_CACHE["section_1"] = s1

    vinculados = []
    for i in range(1, 6):
        vinculados.append({
            "numero": str(i),
            "nombre_oferente": f"Vinculado G5-{i}",
            "cedula": f"{i}" * 10,
            "discapacidad": "Discapacidad física",
            "tipo_contrato": "Contrato temporal",
            "desarrollo_actividad": "Actividad grupal 5 E2E" if i == 1 else "",
        })
    mod.FORM_CACHE["section_2"] = vinculados

    mod.FORM_CACHE["section_6"] = {
        "ajustes_recomendaciones": "Ajustes contratacion grupal 5 E2E",
    }

    mod.FORM_CACHE["section_7"] = [
        {"nombre": "Asistente CG5", "cargo": "Director"},
    ]

    result = mod.export_to_excel(clear_cache=False)
    sid = extract_id_from_result(result)
    CREATED_SHEETS.append(sid)
    sheet = mod.SHEET_NAME

    # Block height=52, first block at 19
    # Vinculado 1: offset 0 -> row 23 for nombre (18+5)
    # Vinculado 2: offset 52 -> row 75
    # Vinculado 3: offset 104 -> row 127
    # Vinculado 4: offset 156 -> row 179
    # Vinculado 5: offset 208 -> row 231
    # Section 6: base=125, extra_blocks=3 (5-2), shift=3*52=156 -> 125+156=281
    # Section 7: base=131, shift=156 -> 131+156=287
    verify_cells("contratacion_grupal_5", sid, sheet, {
        "D7": "2026-03-26",
        "L7": "Mixta",
        "D8": "Empresa ContInc Grup5",
        "L8": "Cartagena",
        # Shared actividad
        "A15": "Actividad grupal 5 E2E",
        # Vinculado 1 nombre
        "C23": "Vinculado G5-1",
        "H23": "1111111111",
        # Vinculado 2 (offset 52)
        "C75": "Vinculado G5-2",
        "H75": "2222222222",
        # Vinculado 3 (offset 104)
        "C127": "Vinculado G5-3",
        "H127": "3333333333",
        # Vinculado 4 (offset 156)
        "C179": "Vinculado G5-4",
        "H179": "4444444444",
        # Vinculado 5 (offset 208)
        "C231": "Vinculado G5-5",
        "H231": "5555555555",
        # Section 6 (ajustes row = 125 + 156 = 281)
        "A281": "Ajustes contratacion grupal 5 E2E",
        # Section 7 (start row = 131 + 156 = 287)
        "C287": "Asistente CG5",
        "K287": "Director",
    })
    return sid


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    tests = [
        ("sensibilizacion", test_sensibilizacion),
        ("presentacion_programa", test_presentacion_programa),
        ("induccion_organizacional", test_induccion_organizacional),
        ("induccion_operativa", test_induccion_operativa),
        ("evaluacion_accesibilidad", test_evaluacion_accesibilidad),
        ("condiciones_vacante", test_condiciones_vacante),
        ("seleccion_incluyente_labs", test_seleccion_incluyente_labs),
        ("seleccion_incluyente_individual", test_seleccion_incluyente_individual),
        ("seleccion_incluyente_grupal", test_seleccion_incluyente_grupal),
        ("contratacion_individual", test_contratacion_individual),
        ("contratacion_grupal_2", test_contratacion_grupal_2),
        ("contratacion_grupal_5", test_contratacion_grupal_5),
    ]

    passed = 0
    failed = 0
    errors = 0

    for name, fn in tests:
        try:
            fn()
            # Check last result
            if RESULTS and RESULTS[-1][1] == "PASS":
                passed += 1
            else:
                failed += 1
        except Exception as exc:
            errors += 1
            print(f"\n  [ERROR] {name}: {exc}")
            traceback.print_exc()
        # Small delay to avoid rate limiting
        time.sleep(1)

    print("\n" + "=" * 60)
    print(f"SUMMARY: {passed} passed, {failed} failed, {errors} errors")
    print(f"Total sheets created: {len(CREATED_SHEETS)}")
    print("=" * 60)

    for name, status, ok, total, errs in RESULTS:
        marker = "✓" if status == "PASS" else "✗"
        print(f"  {marker} {name}: {ok}/{total}")

    if CREATED_SHEETS:
        print(f"\nCreated spreadsheets:")
        for sid in CREATED_SHEETS:
            print(f"  https://docs.google.com/spreadsheets/d/{sid}/edit")
