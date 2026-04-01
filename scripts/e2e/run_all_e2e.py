#!/usr/bin/env python
"""
Runner E2E — Ejecuta todos los formularios uno por uno (secuencial).

Mide el tiempo de finalización de cada formulario (desde que se llama
la función de publicación hasta que se recibe la URL de Google Sheets).

Uso:
    cd <worktree>
    python scripts/e2e/run_all_e2e.py

Opciones:
    --only presentacion_programa condiciones_vacante ...  (correr sólo algunos)
    --stop-on-error                                        (detener al primer error)
"""
import sys
import io
import json
import time
import subprocess
import argparse
from datetime import datetime
from pathlib import Path

# Forzar UTF-8 en stdout/stderr para Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = Path(__file__).resolve().parent
LOGS_DIR = PROJECT_ROOT / "logs"
LOGS_DIR.mkdir(exist_ok=True)

# Orden de ejecución (secuencial)
FORMS = [
    ("presentacion_programa",      "Presentacion del Programa"),
    ("reactivacion_programa",      "Reactivacion del Programa"),
    ("evaluacion_accesibilidad",   "Evaluacion de Accesibilidad"),
    ("condiciones_vacante",        "Condiciones de Vacante"),
    ("seleccion_incluyente",       "Seleccion Incluyente"),
    ("contratacion_incluyente",    "Contratacion Incluyente"),
    ("induccion_organizacional",   "Induccion Organizacional"),
    ("induccion_operativa",        "Induccion Operativa"),
    ("sensibilizacion",            "Sensibilizacion"),
    ("seguimientos",               "Seguimientos"),
]


def _parse_result_from_output(stdout: str) -> dict:
    """Extrae elapsed_s y url de la línea impresa por cada script."""
    result = {"elapsed_s": None, "url": None, "status_line": None}
    for line in stdout.splitlines():
        line = line.strip()
        if line.startswith("[OK]") or line.startswith("[NO_URL]") or line.startswith("[ERROR]"):
            result["status_line"] = line
            # Formato: [STATUS] [form_id] 12.34s -> https://...
            parts = line.split("->")
            if len(parts) == 2:
                url_part = parts[1].strip()
                result["url"] = url_part if url_part.startswith("http") else None
            time_part = parts[0] if len(parts) >= 1 else ""
            for token in time_part.split():
                if token.endswith("s") and token[:-1].replace(".", "").isdigit():
                    try:
                        result["elapsed_s"] = float(token[:-1])
                    except ValueError:
                        pass
    return result


def run_all(only=None, stop_on_error=False):
    started_at = datetime.now()
    print("=" * 80)
    print(f"E2E — CORONA INDUSTRIAL SAS (NIT 900696296-4) — Fecha: {started_at:%Y-%m-%d}")
    print(f"Inicio: {started_at:%Y-%m-%d %H:%M:%S}")
    print("=" * 80)

    results = []

    for form_id, label in FORMS:
        if only and form_id not in only:
            continue

        script_path = SCRIPTS_DIR / f"e2e_{form_id}.py"
        if not script_path.exists():
            print(f"\n[SKIP] {label} ({form_id}) — script no encontrado: {script_path}")
            results.append({
                "form_id": form_id,
                "label": label,
                "status": "SKIP",
                "elapsed_s": None,
                "url": None,
                "error": "Script no encontrado",
                "stdout": "",
                "stderr": "",
            })
            continue

        print(f"\n{'-' * 60}")
        print(f"Formulario: {label} ({form_id})")
        print(f"{'-' * 60}")

        wall_start = time.time()
        proc = subprocess.run(
            [sys.executable, str(script_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=str(PROJECT_ROOT),
        )
        wall_elapsed = time.time() - wall_start

        stdout = proc.stdout or ""
        stderr = proc.stderr or ""

        if stdout:
            print(stdout.rstrip())
        if stderr:
            print("[STDERR]", file=sys.stderr)
            print(stderr.rstrip(), file=sys.stderr)

        parsed = _parse_result_from_output(stdout)
        elapsed_s = parsed["elapsed_s"] if parsed["elapsed_s"] is not None else round(wall_elapsed, 2)
        url = parsed["url"]

        if proc.returncode == 0 and url:
            status = "OK"
        elif proc.returncode == 0 and not url:
            status = "NO_URL"
        else:
            status = "ERROR"

        # Intentar extraer error del stderr o stdout
        error_msg = None
        if status == "ERROR":
            for line in (stderr + "\n" + stdout).splitlines():
                if "[ERROR]" in line or "Error" in line or "Traceback" in line:
                    error_msg = line.strip()
                    break

        results.append({
            "form_id": form_id,
            "label": label,
            "status": status,
            "elapsed_s": elapsed_s,
            "url": url,
            "error": error_msg,
            "returncode": proc.returncode,
        })

        if status == "ERROR" and stop_on_error:
            print(f"\n[STOP] Deteniendo por --stop-on-error después de fallo en {form_id}")
            break

    ended_at = datetime.now()
    total_elapsed = (ended_at - started_at).total_seconds()

    # --- Reporte final ---
    print(f"\n{'=' * 80}")
    print("REPORTE FINAL")
    print(f"{'=' * 80}")
    ok = sum(1 for r in results if r["status"] == "OK")
    fail = sum(1 for r in results if r["status"] in ("ERROR", "NO_URL"))
    skip = sum(1 for r in results if r["status"] == "SKIP")

    col_w = 30
    print(f"{'Formulario':<{col_w}} {'Status':<8} {'Tiempo(s)':<12} URL")
    print(f"{'-'*col_w} {'-'*8} {'-'*12} {'-'*40}")
    for r in results:
        elapsed_str = f"{r['elapsed_s']:.2f}" if r['elapsed_s'] is not None else "—"
        url_str = r["url"] or r.get("error") or "—"
        print(f"{r['label']:<{col_w}} {r['status']:<8} {elapsed_str:<12} {url_str}")

    print(f"{'-' * 80}")
    print(f"OK: {ok} | ERROR/NO_URL: {fail} | SKIP: {skip} | Total: {len(results)}")
    print(f"Tiempo total: {total_elapsed:.1f}s")
    print(f"{'=' * 80}")

    # --- Guardar log JSON ---
    log_filename = LOGS_DIR / f"e2e_{started_at:%Y%m%d_%H%M%S}.json"
    log_data = {
        "started_at": started_at.isoformat(),
        "ended_at": ended_at.isoformat(),
        "total_elapsed_s": round(total_elapsed, 2),
        "empresa": "CORONA INDUSTRIAL SAS",
        "nit": "900696296-4",
        "fecha_visita": "2026-05-05",
        "summary": {"ok": ok, "fail": fail, "skip": skip, "total": len(results)},
        "results": [
            {k: v for k, v in r.items() if k not in ("returncode",)}
            for r in results
        ],
    }
    with open(log_filename, "w", encoding="utf-8") as f:
        json.dump(log_data, f, ensure_ascii=False, indent=2)
    print(f"\nLog guardado en: {log_filename}")

    return results


def main():
    parser = argparse.ArgumentParser(
        description="Runner E2E — corre todos los formularios secuencialmente contra Google Sheets."
    )
    parser.add_argument(
        "--only",
        nargs="*",
        metavar="FORM_ID",
        help="Lista de form_id para ejecutar (e.g. sensibilizacion condiciones_vacante)",
    )
    parser.add_argument(
        "--stop-on-error",
        action="store_true",
        help="Detener al primer formulario con error.",
    )
    args = parser.parse_args()

    only = set(args.only) if args.only else None
    results = run_all(only=only, stop_on_error=args.stop_on_error)
    any_fail = any(r["status"] in ("ERROR", "NO_URL") for r in results)
    sys.exit(1 if any_fail else 0)


if __name__ == "__main__":
    main()
