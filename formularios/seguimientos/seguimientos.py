import os
import re
import shutil

from openpyxl import load_workbook

from formularios.common import (
    _get_desktop_dir,
    _normalize_cedula,
    _sanitize_filename,
    _supabase_get,
)


FORM_ID = "seguimientos"
FORM_NAME = "Seguimientos"
SHARED_ROOT = r"G:\Unidades compartidas\RECA BDs\SEGUIMIENTOS"

SHEET_BASE = "9.  SEGUIMIENTO AL PROCESO DE I"
SHEET_PREFIX = "SEGUIMIENTO PROCESO IL "
SHEET_FINAL = "PONDERADO FINAL"
SHEET_META = "_META_IL"

MODALIDAD_OPTIONS = ["Presencial", "Virtual", "Mixta", "No aplica"]
SI_NO_NA_OPTIONS = ["Si", "No", "No aplica"]
EVAL_OPTIONS = ["Excelente", "Bien", "Necesita mejorar", "Mal", "No aplica"]
TIPO_APOYO_OPTIONS = [
    "Requiere apoyo bajo.",
    "Requiere apoyo medio.",
    "Requiere apoyo Alto.",
    "No requiere apoyo.",
]

SECTION_1_SUPABASE_MAP = {
    "nombre_empresa": "nombre_empresa",
    "ciudad_empresa": "ciudad_empresa",
    "direccion_empresa": "direccion_empresa",
    "nit_empresa": "nit_empresa",
    "correo_1": "correo_1",
    "telefono_empresa": "telefono_empresa",
    "contacto_empresa": "contacto_empresa",
    "cargo": "cargo",
    "asesor": "asesor",
    "sede_empresa": "sede_empresa",
    "caja_compensacion": "caja_compensacion",
}

def register_form():
    return {"id": FORM_ID, "name": FORM_NAME, "module": __name__}


def _ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def _get_local_root():
    desktop = _get_desktop_dir()
    return _ensure_dir(os.path.join(desktop, "Formatos Inclusion Laboral", "SEGUIMIENTOS"))


def _get_roots():
    local_root = _get_local_root()
    shared_root = SHARED_ROOT
    try:
        _ensure_dir(shared_root)
        shared_ok = True
    except Exception:
        shared_ok = False
    return {
        "local": local_root,
        "shared": shared_root if shared_ok else None,
    }


def _find_template_path():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    templates_dir = os.path.join(base_dir, "templates")
    if not os.path.isdir(templates_dir):
        raise FileNotFoundError("No existe la carpeta templates.")
    for name in os.listdir(templates_dir):
        if name.startswith("~$"):
            continue
        if name.lower().endswith(".xlsx") and "seguimiento" in name.lower():
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de seguimientos.")


def _parse_first_name_lastname(full_name):
    tokens = [t for t in re.split(r"\s+", str(full_name or "").strip()) if t]
    if not tokens:
        return "Usuario", "SinApellido"
    first_name = tokens[0]
    if len(tokens) >= 4:
        first_lastname = tokens[2]
    elif len(tokens) >= 2:
        first_lastname = tokens[1]
    else:
        first_lastname = "SinApellido"
    return first_name, first_lastname


def build_case_folder_name(nombre_usuario, cedula):
    first_name, first_lastname = _parse_first_name_lastname(nombre_usuario)
    base = f"{first_name} {first_lastname} - {cedula}"
    return _sanitize_filename(base)


def _find_excel_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        return None
    for name in os.listdir(folder_path):
        if name.startswith("~$"):
            continue
        if name.lower().endswith(".xlsx"):
            return os.path.join(folder_path, name)
    return None


def _scan_case_folder_by_cedula(root, cedula):
    if not root or not os.path.isdir(root):
        return None
    suffix = f"- {cedula}"
    for name in os.listdir(root):
        full = os.path.join(root, name)
        if not os.path.isdir(full):
            continue
        if name.endswith(suffix):
            found = _find_excel_in_folder(full)
            if found:
                return found
    return None


def find_case_workbook(cedula, nombre_usuario=""):
    normalized = _normalize_cedula(cedula)
    if not normalized:
        return None
    roots = _get_roots()
    folder_name = build_case_folder_name(nombre_usuario, normalized)
    ordered_roots = [roots.get("shared"), roots.get("local")]
    for root in ordered_roots:
        if not root:
            continue
        direct = _find_excel_in_folder(os.path.join(root, folder_name))
        if direct:
            return direct
        scanned = _scan_case_folder_by_cedula(root, normalized)
        if scanned:
            return scanned
    return None


def _get_str(value):
    if value is None:
        return ""
    return str(value).strip()


def _set_if_empty(ws, cell, value):
    if ws[cell].value in (None, "") and value not in (None, ""):
        ws[cell].value = value


def _fill_sheet_base(wb, user_row):
    if SHEET_BASE not in wb.sheetnames:
        return
    ws = wb[SHEET_BASE]
    _set_if_empty(ws, "A16", _get_str(user_row.get("nombre_usuario")))
    _set_if_empty(ws, "E16", _get_str(user_row.get("cedula_usuario")))
    _set_if_empty(ws, "I16", _get_str(user_row.get("telefono_oferente")))
    _set_if_empty(ws, "K16", _get_str(user_row.get("correo_oferente")))
    _set_if_empty(ws, "P16", _get_str(user_row.get("contacto_emergencia")))
    _set_if_empty(ws, "S16", _get_str(user_row.get("parentesco")))
    _set_if_empty(ws, "U16", _get_str(user_row.get("telefono_emergencia")))
    _set_if_empty(ws, "A18", _get_str(user_row.get("cargo_oferente")))
    _set_if_empty(ws, "I18", _get_str(user_row.get("certificado_porcentaje")))
    discapacidad = _get_str(user_row.get("discapacidad_detalle")) or _get_str(
        user_row.get("discapacidad_usuario")
    )
    _set_if_empty(ws, "N18", discapacidad)


def _ensure_meta_sheet(wb, cedula, nombre_usuario, is_compensar, max_seguimientos):
    if SHEET_META in wb.sheetnames:
        ws = wb[SHEET_META]
    else:
        ws = wb.create_sheet(SHEET_META)
    ws.sheet_state = "hidden"
    ws["A1"] = "cedula"
    ws["B1"] = _get_str(cedula)
    ws["A2"] = "nombre_usuario"
    ws["B2"] = _get_str(nombre_usuario)
    ws["A3"] = "is_compensar"
    ws["B3"] = "1" if is_compensar else "0"
    ws["A4"] = "max_seguimientos"
    ws["B4"] = int(max_seguimientos)


def _read_meta(wb):
    if SHEET_META not in wb.sheetnames:
        return {}
    ws = wb[SHEET_META]
    meta = {}
    for row in range(1, 15):
        key = _get_str(ws[f"A{row}"].value)
        val = ws[f"B{row}"].value
        if key:
            meta[key] = val
    return meta


def _apply_visibility(wb, max_seguimientos):
    limit = 6 if int(max_seguimientos or 0) >= 6 else 3
    for i in range(1, 7):
        name = f"{SHEET_PREFIX}{i}"
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        ws.sheet_state = "visible" if i <= limit else "hidden"


def _copy_to_secondary_roots(primary_path):
    roots = _get_roots()
    local = roots.get("local")
    shared = roots.get("shared")
    if not primary_path or not os.path.exists(primary_path):
        return
    folder_name = os.path.basename(os.path.dirname(primary_path))
    filename = os.path.basename(primary_path)
    targets = []
    if shared and not os.path.normcase(primary_path).startswith(os.path.normcase(shared)):
        targets.append(os.path.join(shared, folder_name, filename))
    if local and not os.path.normcase(primary_path).startswith(os.path.normcase(local)):
        targets.append(os.path.join(local, folder_name, filename))
    for target in targets:
        try:
            _ensure_dir(os.path.dirname(target))
            shutil.copy2(primary_path, target)
        except Exception:
            continue


def ensure_case_workbook(cedula, user_row, is_compensar):
    normalized = _normalize_cedula(cedula)
    if not normalized:
        raise ValueError("Cédula inválida.")
    if not user_row:
        raise ValueError("No se encontró usuario para la cédula indicada.")

    existing = find_case_workbook(normalized, user_row.get("nombre_usuario"))
    if existing:
        wb = load_workbook(existing)
        meta = _read_meta(wb)
        if meta:
            max_seguimientos = int(meta.get("max_seguimientos") or (6 if is_compensar else 3))
            is_comp = str(meta.get("is_compensar") or "0").strip() in ("1", "true", "True")
        else:
            is_comp = bool(is_compensar)
            max_seguimientos = 6 if is_comp else 3
        _fill_sheet_base(wb, user_row)
        _ensure_meta_sheet(
            wb,
            normalized,
            user_row.get("nombre_usuario"),
            is_comp,
            max_seguimientos,
        )
        _apply_visibility(wb, max_seguimientos)
        wb.save(existing)
        _copy_to_secondary_roots(existing)
        return {"path": existing, "created": False, "max_seguimientos": max_seguimientos}

    template_path = _find_template_path()
    roots = _get_roots()
    primary_root = roots.get("shared") or roots.get("local")
    if not primary_root:
        raise RuntimeError("No hay ruta disponible para guardar seguimientos.")

    folder_name = build_case_folder_name(user_row.get("nombre_usuario"), normalized)
    case_folder = _ensure_dir(os.path.join(primary_root, folder_name))
    output_path = os.path.join(case_folder, f"{folder_name}.xlsx")
    shutil.copy2(template_path, output_path)

    max_seguimientos = 6 if bool(is_compensar) else 3
    wb = load_workbook(output_path)
    _fill_sheet_base(wb, user_row)
    _ensure_meta_sheet(
        wb,
        normalized,
        user_row.get("nombre_usuario"),
        bool(is_compensar),
        max_seguimientos,
    )
    _apply_visibility(wb, max_seguimientos)
    wb.save(output_path)
    _copy_to_secondary_roots(output_path)
    return {"path": output_path, "created": True, "max_seguimientos": max_seguimientos}


def _cell_has_value(ws, cell):
    value = ws[cell].value
    if value is None:
        return False
    return str(value).strip() != ""


def _is_base_completed(wb):
    if SHEET_BASE not in wb.sheetnames:
        return False
    ws = wb[SHEET_BASE]
    required = ["A16", "E16", "A18", "N18"]
    return all(_cell_has_value(ws, c) for c in required)


def _is_followup_completed(ws):
    required = ["O12", "R12", "J31"]
    return all(_cell_has_value(ws, c) for c in required)


def suggest_next_step(workbook_path):
    if not workbook_path or not os.path.exists(workbook_path):
        return {"sheet": SHEET_BASE, "message": "Inicia con la hoja base."}
    wb = load_workbook(workbook_path, data_only=False)
    meta = _read_meta(wb)
    try:
        max_seguimientos = int(meta.get("max_seguimientos") or 6)
    except Exception:
        max_seguimientos = 6
    max_seguimientos = 6 if max_seguimientos >= 6 else 3

    if not _is_base_completed(wb):
        return {
            "sheet": SHEET_BASE,
            "message": "Completa primero la hoja base del proceso.",
            "max_seguimientos": max_seguimientos,
        }

    for i in range(1, max_seguimientos + 1):
        name = f"{SHEET_PREFIX}{i}"
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        if not _is_followup_completed(ws):
            return {
                "sheet": name,
                "message": f"Siguiente sugerido: seguimiento {i}.",
                "max_seguimientos": max_seguimientos,
            }

    return {
        "sheet": SHEET_FINAL,
        "message": "Todos los seguimientos están diligenciados. Revisa el ponderado final.",
        "max_seguimientos": max_seguimientos,
    }


def get_usuarios_reca_cedulas(env_path=".env"):
    params = {
        "select": "cedula_usuario",
        "cedula_usuario": "not.is.null",
        "order": "cedula_usuario.asc",
    }
    data = _supabase_get("usuarios_reca", params, env_path=env_path)
    return [row.get("cedula_usuario") for row in data if row.get("cedula_usuario")]


def get_usuario_reca_by_cedula(cedula, env_path=".env"):
    normalized = _normalize_cedula(cedula)
    if not normalized:
        return None
    select_cols = ",".join(
        [
            "cedula_usuario",
            "nombre_usuario",
            "discapacidad_usuario",
            "discapacidad_detalle",
            "certificado_porcentaje",
            "telefono_oferente",
            "correo_oferente",
            "cargo_oferente",
            "contacto_emergencia",
            "parentesco",
            "telefono_emergencia",
        ]
    )
    params = {
        "select": select_cols,
        "cedula_usuario": f"eq.{normalized}",
        "limit": 1,
    }
    data = _supabase_get("usuarios_reca", params, env_path=env_path)
    return data[0] if data else None


def get_empresa_by_nit(nit, env_path=".env"):
    if not nit:
        return None
    nit = "".join(str(nit).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nit_empresa": f"eq.{nit}",
        "limit": 1,
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    return data[0] if data else None


def get_empresa_by_nombre(nombre, env_path=".env"):
    if not nombre:
        return None
    nombre = " ".join(str(nombre).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nombre_empresa": f"ilike.{nombre}",
        "limit": 2,
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    if not data:
        return None
    if len(data) > 1:
        raise ValueError("Hay más de una empresa con ese nombre. Usa el NIT.")
    return data[0]


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=10):
    text = str(prefix or "").strip()
    if not text:
        return []
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nombre_empresa": f"ilike.{text}%",
        "order": "nombre_empresa.asc",
        "limit": int(limit),
    }
    return _supabase_get("empresas", params, env_path=env_path)


def _ensure_sheet_exists(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"No existe la hoja '{sheet_name}' en el archivo.")
    return wb[sheet_name]


def _cell_value(ws, address):
    value = ws[address].value
    if value is None:
        return ""
    return str(value).strip()


def get_case_meta(workbook_path):
    wb = load_workbook(workbook_path, data_only=False)
    meta = _read_meta(wb)
    try:
        max_seg = int(meta.get("max_seguimientos") or 6)
    except Exception:
        max_seg = 6
    max_seg = 6 if max_seg >= 6 else 3
    return {
        "cedula": _get_str(meta.get("cedula")),
        "nombre_usuario": _get_str(meta.get("nombre_usuario")),
        "is_compensar": str(meta.get("is_compensar") or "0").strip() in ("1", "true", "True"),
        "max_seguimientos": max_seg,
    }


def get_base_payload(workbook_path):
    wb = load_workbook(workbook_path, data_only=False)
    ws = _ensure_sheet_exists(wb, SHEET_BASE)
    payload = {
        "fecha_visita": _cell_value(ws, "D8"),
        "modalidad": _cell_value(ws, "R8"),
        "nombre_empresa": _cell_value(ws, "D9"),
        "ciudad_empresa": _cell_value(ws, "R9"),
        "direccion_empresa": _cell_value(ws, "D10"),
        "nit_empresa": _cell_value(ws, "R10"),
        "correo_1": _cell_value(ws, "D11"),
        "telefono_empresa": _cell_value(ws, "R11"),
        "contacto_empresa": _cell_value(ws, "D12"),
        "cargo": _cell_value(ws, "R12"),
        "asesor": _cell_value(ws, "D13"),
        "sede_empresa": _cell_value(ws, "R13"),
        "nombre_vinculado": _cell_value(ws, "A16"),
        "cedula": _cell_value(ws, "E16"),
        "telefono_vinculado": _cell_value(ws, "I16"),
        "correo_vinculado": _cell_value(ws, "K16"),
        "contacto_emergencia": _cell_value(ws, "P16"),
        "parentesco": _cell_value(ws, "S16"),
        "telefono_emergencia": _cell_value(ws, "U16"),
        "cargo_vinculado": _cell_value(ws, "A18"),
        "certificado_discapacidad": _cell_value(ws, "E18"),
        "certificado_porcentaje": _cell_value(ws, "I18"),
        "discapacidad": _cell_value(ws, "N18"),
        "tipo_contrato": _cell_value(ws, "C20"),
        "fecha_inicio_contrato": _cell_value(ws, "M20"),
        "fecha_fin_contrato": _cell_value(ws, "T20"),
        "apoyos_ajustes": _cell_value(ws, "A21"),
        "funciones_1_5": [_cell_value(ws, f"B{r}") for r in range(23, 28)],
        "funciones_6_10": [_cell_value(ws, f"N{r}") for r in range(23, 28)],
        "seguimiento_fechas_1_3": [_cell_value(ws, f"C{r}") for r in range(29, 32)],
        "seguimiento_fechas_4_6": [_cell_value(ws, f"P{r}") for r in range(29, 32)],
    }
    return payload


def save_base_payload(workbook_path, payload):
    wb = load_workbook(workbook_path, data_only=False)
    ws = _ensure_sheet_exists(wb, SHEET_BASE)
    mapping = {
        "fecha_visita": "D8",
        "modalidad": "R8",
        "nombre_empresa": "D9",
        "ciudad_empresa": "R9",
        "direccion_empresa": "D10",
        "nit_empresa": "R10",
        "correo_1": "D11",
        "telefono_empresa": "R11",
        "contacto_empresa": "D12",
        "cargo": "R12",
        "asesor": "D13",
        "sede_empresa": "R13",
        "nombre_vinculado": "A16",
        "cedula": "E16",
        "telefono_vinculado": "I16",
        "correo_vinculado": "K16",
        "contacto_emergencia": "P16",
        "parentesco": "S16",
        "telefono_emergencia": "U16",
        "cargo_vinculado": "A18",
        "certificado_discapacidad": "E18",
        "certificado_porcentaje": "I18",
        "discapacidad": "N18",
        "tipo_contrato": "C20",
        "fecha_inicio_contrato": "M20",
        "fecha_fin_contrato": "T20",
        "apoyos_ajustes": "A21",
    }
    for key, cell in mapping.items():
        if key in payload:
            ws[cell].value = payload.get(key, "")
    f1 = payload.get("funciones_1_5") or []
    f2 = payload.get("funciones_6_10") or []
    for i, row in enumerate(range(23, 28)):
        if i < len(f1):
            ws[f"B{row}"].value = f1[i]
        if i < len(f2):
            ws[f"N{row}"].value = f2[i]
    s1 = payload.get("seguimiento_fechas_1_3") or []
    s2 = payload.get("seguimiento_fechas_4_6") or []
    for i, row in enumerate(range(29, 32)):
        if i < len(s1):
            ws[f"C{row}"].value = s1[i]
        if i < len(s2):
            ws[f"P{row}"].value = s2[i]
    wb.save(workbook_path)


def _get_followup_sheet_name(index):
    idx = int(index)
    if idx < 1 or idx > 6:
        raise ValueError("El seguimiento debe estar entre 1 y 6.")
    return f"{SHEET_PREFIX}{idx}"


def get_followup_payload(workbook_path, index):
    wb = load_workbook(workbook_path, data_only=False)
    ws = _ensure_sheet_exists(wb, _get_followup_sheet_name(index))
    item_labels = [_cell_value(ws, f"A{r}") for r in range(12, 31)]
    empresa_labels = [_cell_value(ws, f"A{r}") for r in range(34, 42)]
    payload = {
        "modalidad": _cell_value(ws, "E8"),
        "seguimiento_numero": _cell_value(ws, "P8"),
        "item_labels": item_labels,
        "item_observaciones": [_cell_value(ws, f"G{r}") for r in range(12, 31)],
        "item_autoevaluacion": [_cell_value(ws, f"O{r}") for r in range(12, 31)],
        "item_eval_empresa": [_cell_value(ws, f"R{r}") for r in range(12, 31)],
        "tipo_apoyo": _cell_value(ws, "J31"),
        "empresa_item_labels": empresa_labels,
        "empresa_eval": [_cell_value(ws, f"J{r}") for r in range(34, 42)],
        "empresa_observacion": [_cell_value(ws, f"L{r}") for r in range(34, 42)],
        "situacion_encontrada": _cell_value(ws, "A43"),
        "estrategias_ajustes": _cell_value(ws, "A45"),
        "asistentes": [
            {"nombre": _cell_value(ws, f"D{r}"), "cargo": _cell_value(ws, f"N{r}")}
            for r in range(47, 51)
        ],
    }
    return payload


def save_followup_payload(workbook_path, index, payload):
    wb = load_workbook(workbook_path, data_only=False)
    ws = _ensure_sheet_exists(wb, _get_followup_sheet_name(index))
    ws["E8"].value = payload.get("modalidad", "")
    ws["P8"].value = payload.get("seguimiento_numero", index)
    item_obs = payload.get("item_observaciones") or []
    item_auto = payload.get("item_autoevaluacion") or []
    item_emp = payload.get("item_eval_empresa") or []
    for i, row in enumerate(range(12, 31)):
        if i < len(item_obs):
            ws[f"G{row}"].value = item_obs[i]
        if i < len(item_auto):
            ws[f"O{row}"].value = item_auto[i]
        if i < len(item_emp):
            ws[f"R{row}"].value = item_emp[i]
    ws["J31"].value = payload.get("tipo_apoyo", "")
    emp_eval = payload.get("empresa_eval") or []
    emp_obs = payload.get("empresa_observacion") or []
    for i, row in enumerate(range(34, 42)):
        if i < len(emp_eval):
            ws[f"J{row}"].value = emp_eval[i]
        if i < len(emp_obs):
            ws[f"L{row}"].value = emp_obs[i]
    ws["A43"].value = payload.get("situacion_encontrada", "")
    ws["A45"].value = payload.get("estrategias_ajustes", "")
    asistentes = payload.get("asistentes") or []
    for i, row in enumerate(range(47, 51)):
        entry = asistentes[i] if i < len(asistentes) else {}
        ws[f"D{row}"].value = entry.get("nombre", "")
        ws[f"N{row}"].value = entry.get("cargo", "")
    wb.save(workbook_path)
