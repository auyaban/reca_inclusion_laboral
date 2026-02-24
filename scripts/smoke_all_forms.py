import argparse
import inspect
import os
import random
import string
import sys
import traceback
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import drive_upload
from formularios.presentacion_programa import presentacion_programa
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.condiciones_vacante import condiciones_vacante
from formularios.seleccion_incluyente import seleccion_incluyente
from formularios.contratacion_incluyente import contratacion_incluyente
from formularios.induccion_organizacional import induccion_organizacional
from formularios.induccion_operativa import induccion_operativa
from formularios.sensibilizacion import sensibilizacion


FORMS = [
    ("presentacion_programa", "Presentacion Programa", presentacion_programa),
    ("evaluacion_accesibilidad", "Evaluacion Accesibilidad", evaluacion_accesibilidad),
    ("condiciones_vacante", "Revision Condicion", condiciones_vacante),
    ("seleccion_incluyente", "Seleccion Incluyente", seleccion_incluyente),
    ("contratacion_incluyente", "Contratacion Incluyente", contratacion_incluyente),
    ("induccion_organizacional", "Induccion Organizacional", induccion_organizacional),
    ("induccion_operativa", "Induccion Operativa", induccion_operativa),
    ("sensibilizacion", "Sensibilizacion", sensibilizacion),
]


def _rand_digits(n=10):
    return "".join(random.choice(string.digits) for _ in range(n))


def _rand_word(prefix="TEST"):
    return f"{prefix}_{random.randint(1000, 9999)}"


def _random_value_for_field(field_id):
    fid = str(field_id or "").lower()
    if "fecha" in fid:
        return datetime.now().strftime("%Y-%m-%d")
    if "correo" in fid:
        return f"{_rand_word('mail').lower()}@example.com"
    if "telefono" in fid or "cel" in fid:
        return _rand_digits(10)
    if "nit" in fid:
        return f"{_rand_digits(9)}-{random.randint(0,9)}"
    if "modalidad" in fid:
        return random.choice(["Virtual", "Presencial", "Mixto"])
    if "tipo_visita" in fid:
        return random.choice(["Presentacion", "Reactivacion"])
    if "ciudad" in fid or "municipio" in fid:
        return random.choice(["Bogota", "Medellin", "Cali", "Barranquilla"])
    if "nombre" in fid and "empresa" in fid:
        return f"EMPRESA {_rand_word('SMOKE')}"
    if "empresa" in fid:
        return f"EMPRESA {_rand_word('SMOKE')}"
    if "sede" in fid:
        return "Principal"
    if "caja" in fid:
        return "Compensar"
    if "asesor" in fid or "profesional" in fid or "contacto" in fid or "cargo" in fid:
        return "Usuario Prueba"
    return f"VAL_{_rand_word('S1')}"


def _build_random_section_1(module, form_label):
    payload = {}
    section1 = getattr(module, "SECTION_1", {}) or {}
    fields = section1.get("fields", []) if isinstance(section1, dict) else []
    for field in fields:
        if not isinstance(field, dict):
            continue
        field_id = field.get("id")
        if not field_id:
            continue
        payload[field_id] = _random_value_for_field(field_id)

    payload.setdefault("nombre_empresa", f"EMPRESA {_rand_word('SMOKE')}")
    payload.setdefault("fecha_visita", datetime.now().strftime("%Y-%m-%d"))
    payload.setdefault("modalidad", "Virtual")
    payload.setdefault("nit_empresa", f"{_rand_digits(9)}-{random.randint(0,9)}")
    if "presentacion" in form_label.lower():
        payload.setdefault("tipo_visita", "Presentacion")
    return payload


def _seed_random_cache(module, form_label):
    if hasattr(module, "clear_cache_file"):
        module.clear_cache_file()
    if hasattr(module, "clear_form_cache"):
        module.clear_form_cache()

    section_1_payload = _build_random_section_1(module, form_label)
    if hasattr(module, "confirm_section_1"):
        # Use canonical entrypoint so module-level SECTION_1_CACHE is populated too.
        module.confirm_section_1(section_1_payload, section_1_payload)
    elif hasattr(module, "set_section_cache"):
        module.set_section_cache("section_1", section_1_payload)
    else:
        raise RuntimeError("El modulo no expone confirm_section_1 ni set_section_cache")

    if hasattr(module, "save_cache_to_file"):
        module.save_cache_to_file()

    return section_1_payload


def _export_module(module):
    export_fn = getattr(module, "export_to_excel", None)
    if export_fn is None:
        raise RuntimeError("No tiene export_to_excel")

    sig = inspect.signature(export_fn)
    kwargs = {}
    if "clear_cache" in sig.parameters:
        kwargs["clear_cache"] = False
    if "cache" in sig.parameters and hasattr(module, "get_form_cache"):
        kwargs["cache"] = module.get_form_cache()
    if "progress_callback" in sig.parameters:
        kwargs["progress_callback"] = None
    return export_fn(**kwargs)


def run_smoke(only=None, strict=False, with_drive=False, professional="Smoke Test"):
    started = datetime.now()
    results = []

    for form_id, label, module in FORMS:
        if only and form_id not in only:
            continue

        try:
            section_1 = _seed_random_cache(module, label)
            output_path = _export_module(module)
            if not output_path:
                raise RuntimeError("export_to_excel no devolvio ruta")

            drive_result = None
            if with_drive:
                file_id, file_name = drive_upload.upload_excel_to_drive(
                    output_path,
                    base_name=os.path.basename(output_path),
                    professional_name=professional,
                )
                drive_result = f"OK id={file_id} name={file_name}"

            results.append(
                {
                    "form_id": form_id,
                    "label": label,
                    "status": "OK",
                    "detail": "Exportado con datos random",
                    "empresa": section_1.get("nombre_empresa"),
                    "excel": output_path,
                    "drive": drive_result,
                }
            )
        except Exception as exc:
            detail = f"{exc}\n{traceback.format_exc(limit=1)}".strip()
            results.append(
                {
                    "form_id": form_id,
                    "label": label,
                    "status": "FAIL",
                    "detail": detail,
                    "empresa": None,
                    "excel": None,
                    "drive": None,
                }
            )
            if strict:
                break

    ended = datetime.now()
    return started, ended, results


def print_report(started, ended, results):
    print("=" * 100)
    print("SMOKE TEST - EXPORT CON DATOS RANDOM (SIN CACHE REAL)")
    print(f"Inicio: {started:%Y-%m-%d %H:%M:%S}")
    print(f"Fin:    {ended:%Y-%m-%d %H:%M:%S}")
    print("=" * 100)

    ok = fail = 0
    for row in results:
        status = row["status"]
        if status == "OK":
            ok += 1
        else:
            fail += 1

        print(f"[{status}] {row['label']} ({row['form_id']})")
        print(f"  Detalle: {row['detail']}")
        if row.get("empresa"):
            print(f"  Empresa random: {row['empresa']}")
        if row.get("excel"):
            print(f"  Excel: {row['excel']}")
        if row.get("drive"):
            print(f"  Drive: {row['drive']}")

    print("-" * 100)
    print(f"Resumen -> OK: {ok} | FAIL: {fail} | Total evaluados: {len(results)}")
    print("=" * 100)


def main():
    parser = argparse.ArgumentParser(
        description="Smoke test de formularios usando datos random (sin datos reales)."
    )
    parser.add_argument(
        "--only",
        nargs="*",
        help="Lista de form_id para ejecutar (ej: seleccion_incluyente contratacion_incluyente)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Detener al primer FAIL.",
    )
    parser.add_argument(
        "--with-drive",
        action="store_true",
        help="Sube cada Excel generado a Drive.",
    )
    parser.add_argument(
        "--professional",
        default="Smoke Test",
        help="Nombre del profesional para carpeta de Drive.",
    )
    args = parser.parse_args()

    only = set(args.only) if args.only else None
    started, ended, results = run_smoke(
        only=only,
        strict=args.strict,
        with_drive=args.with_drive,
        professional=args.professional,
    )
    print_report(started, ended, results)


if __name__ == "__main__":
    main()
