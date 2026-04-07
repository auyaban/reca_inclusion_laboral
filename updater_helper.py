from __future__ import annotations

import argparse
import sys

from updater import execute_update_manifest, mark_update_failed_from_manifest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="RECA updater helper")
    parser.add_argument("--manifest", required=True, help="Ruta al manifest de actualización")
    args = parser.parse_args(argv)
    try:
        result = execute_update_manifest(args.manifest)
    except Exception as exc:
        mark_update_failed_from_manifest(args.manifest, str(exc))
        return 1
    return 0 if str(result.get("state") or "").strip().lower() == "succeeded" else 1


if __name__ == "__main__":
    raise SystemExit(main())
