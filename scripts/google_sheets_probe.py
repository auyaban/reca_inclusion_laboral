import argparse
import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from google_sheets_client import (
    get_default_spreadsheet_id,
    get_spreadsheet,
    read_sheet_values,
    write_sheet_values,
)


def _build_parser():
    parser = argparse.ArgumentParser(
        description="Prueba segura de Google Sheets con service account."
    )
    parser.add_argument(
        "--spreadsheet",
        help=(
            "Spreadsheet ID o URL completa. "
            "Si se omite, usa GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID."
        ),
    )
    parser.add_argument(
        "--range",
        help="Rango A1 para lectura, por ejemplo Hoja1!A1:D10.",
    )
    parser.add_argument(
        "--write-range",
        help="Rango A1 para escritura. Requiere --write-json con matriz 2D.",
    )
    parser.add_argument(
        "--write-json",
        help='JSON de matriz 2D para escritura, por ejemplo [["hola","mundo"]].',
    )
    return parser


def main():
    args = _build_parser().parse_args()
    spreadsheet = args.spreadsheet or get_default_spreadsheet_id()

    meta = get_spreadsheet(spreadsheet, include_grid_data=False)
    print(
        json.dumps(
            {
                "spreadsheetId": meta.get("spreadsheetId"),
                "title": meta.get("properties", {}).get("title"),
                "sheets": [
                    sheet.get("properties", {}).get("title")
                    for sheet in meta.get("sheets", [])
                ],
            },
            ensure_ascii=False,
            indent=2,
        )
    )

    if args.range:
        values = read_sheet_values(spreadsheet, args.range)
        print(json.dumps({"range": args.range, "values": values}, ensure_ascii=False, indent=2))

    if args.write_range:
        if not args.write_json:
            raise SystemExit("--write-range requiere --write-json.")
        values = json.loads(args.write_json)
        result = write_sheet_values(spreadsheet, args.write_range, values)
        print(json.dumps({"write_result": result}, ensure_ascii=False, indent=2))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
