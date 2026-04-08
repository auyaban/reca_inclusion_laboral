import unittest
from unittest.mock import patch

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.presentacion_programa import presentacion_programa


class CompanyNameLookupPagingTests(unittest.TestCase):
    def test_evaluacion_lookup_pages_all_matches_when_limit_is_zero(self) -> None:
        rows = [
            {"nombre_empresa": "Alpha SAS"},
            {"nombre_empresa": "Alpine SAS"},
            {"nombre_empresa": "Alpha SAS"},
        ]

        with patch.object(evaluacion_accesibilidad, "_supabase_get_paged", return_value=rows) as paged:
            result = evaluacion_accesibilidad.get_empresas_by_nombre_prefix("Al", limit=0)

        self.assertEqual(result, ["Alpha SAS", "Alpine SAS"])
        paged.assert_called_once()
        self.assertEqual(paged.call_args.args[0], "empresas")
        self.assertEqual(
            paged.call_args.args[1],
            {"select": "nombre_empresa", "nombre_empresa": "ilike.Al%"},
        )

    def test_presentacion_lookup_pages_requested_batches_above_single_page(self) -> None:
        rows = [{"nombre_empresa": f"Empresa {index:03d}"} for index in range(1, 505)]

        with patch.object(presentacion_programa, "_supabase_get_paged", return_value=rows) as paged:
            result = presentacion_programa.get_empresas_by_nombre_prefix("Emp", limit=503)

        self.assertEqual(len(result), 503)
        self.assertEqual(result[0], "Empresa 001")
        self.assertEqual(result[-1], "Empresa 503")
        paged.assert_called_once()
        self.assertEqual(paged.call_args.kwargs["max_pages"], 2)


if __name__ == "__main__":
    unittest.main()
