import unittest
from unittest import mock

import text_review


class _FakeHttpResponse:
    def __init__(self, payload):
        self._payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self):
        return self._payload


class TextReviewTargetTests(unittest.TestCase):
    def test_condiciones_vacante_labs_reuses_regular_review_targets(self) -> None:
        cache_snapshot = {
            "section_2": {
                "requiere_certificado_observaciones": "texto de certificado",
            },
            "section_2_1": {
                "observaciones": "observacion general",
                "funciones_tareas": "funcion uno",
            },
            "section_5": {
                "observaciones_peligros": "observaciones de salud ocupacional",
            },
            "section_7": {
                "observaciones_recomendaciones": "texto libre final",
            },
        }

        targets = text_review.extract_review_targets("condiciones_vacante_labs", cache_snapshot)
        paths = {tuple(item["path"]) for item in targets}

        self.assertIn(("section_2", "requiere_certificado_observaciones"), paths)
        self.assertIn(("section_2_1", "observaciones"), paths)
        self.assertIn(("section_2_1", "funciones_tareas"), paths)
        self.assertIn(("section_5", "observaciones_peligros"), paths)
        self.assertIn(("section_7", "observaciones_recomendaciones"), paths)


class TextReviewBatchTests(unittest.TestCase):
    def test_call_edge_review_requires_valid_session_without_mojibake(self) -> None:
        settings = {"timeout": 30, "function_name": "text-review-orthography"}

        with (
            mock.patch.object(text_review, "_load_supabase_credentials", return_value=("url", "key")),
            mock.patch.object(text_review, "_supabase_get_access_token", return_value=""),
        ):
            with self.assertRaisesRegex(RuntimeError, "No hay sesión válida para revisar ortografía."):
                text_review._call_edge_review({"items": []}, settings)

    def test_call_edge_review_uses_utf8_default_error_message(self) -> None:
        settings = {"timeout": 30, "function_name": "text-review-orthography"}

        with (
            mock.patch.object(text_review, "_load_supabase_credentials", return_value=("https://example.supabase.co", "key")),
            mock.patch.object(text_review, "_supabase_get_access_token", return_value="jwt-token"),
            mock.patch.object(
                text_review.urllib.request,
                "urlopen",
                return_value=_FakeHttpResponse(b'{"ok": false}'),
            ),
        ):
            with self.assertRaisesRegex(RuntimeError, "La función de revisión no devolvió texto."):
                text_review._call_edge_review({"items": []}, settings)

    def test_build_review_batches_respects_item_and_char_limits(self) -> None:
        settings = {
            "batch_max_items": 2,
            "batch_max_chars": 10,
        }

        batches = text_review._build_review_batches(
            ["abcd", "ef", "ghij", "k"],
            settings,
        )

        self.assertEqual(batches, [["abcd", "ef"], ["ghij", "k"]])

    def test_parse_batch_review_output_accepts_json_code_fences(self) -> None:
        reviewed = text_review._parse_batch_review_output(
            """```json
{"items":[{"id":"item_1","text":"Hola"},{"id":"item_2","text":"Mundo"}]}
```""",
            ["item_1", "item_2"],
        )

        self.assertEqual(reviewed, ["Hola", "Mundo"])

    def test_review_export_cache_batches_unique_texts(self) -> None:
        cache_snapshot = {
            "section_2_1": {
                "observaciones": "texto a",
                "funciones_tareas": "texto b",
                "conocimientos_basicos": "texto a",
            }
        }
        settings = {
            "enabled": True,
            "api_key": "",
            "model": "gpt-4.1-nano",
            "function_name": "text-review-orthography",
            "timeout": 45,
            "batch_max_items": 8,
            "batch_max_chars": 12000,
        }

        with (
            mock.patch.object(text_review, "_read_settings", return_value=settings),
            mock.patch.object(text_review, "_load_supabase_credentials", return_value=("url", "key")),
            mock.patch.object(text_review, "_supabase_get_access_token", return_value="jwt"),
            mock.patch.object(
                text_review,
                "_review_text_batch",
                return_value=["texto a corregido", "texto b corregido"],
            ) as batch_mock,
            mock.patch.object(text_review, "_review_text") as single_mock,
        ):
            result = text_review.review_export_cache("evaluacion_accesibilidad", cache_snapshot)

        self.assertEqual(result.status, "reviewed")
        self.assertEqual(result.reviewed_count, 3)
        self.assertEqual(result.cache["section_2_1"]["observaciones"], "texto a corregido")
        self.assertEqual(result.cache["section_2_1"]["conocimientos_basicos"], "texto a corregido")
        self.assertEqual(result.cache["section_2_1"]["funciones_tareas"], "texto b corregido")
        batch_mock.assert_called_once_with(["texto a", "texto b"], settings)
        single_mock.assert_not_called()

    def test_review_text_batch_falls_back_to_individual_reviews(self) -> None:
        settings = {
            "enabled": True,
            "api_key": "",
            "model": "gpt-4.1-nano",
            "function_name": "text-review-orthography",
            "timeout": 45,
            "batch_max_items": 8,
            "batch_max_chars": 12000,
        }

        with (
            mock.patch.object(
                text_review,
                "_review_text_batch_via_edge",
                side_effect=RuntimeError("bad batch"),
            ),
            mock.patch.object(
                text_review,
                "_review_text",
                side_effect=lambda text, _settings: f"{text} corregido",
            ) as single_mock,
        ):
            reviewed = text_review._review_text_batch(["uno", "dos"], settings)

        self.assertEqual(reviewed, ["uno corregido", "dos corregido"])
        self.assertEqual(single_mock.call_count, 2)


if __name__ == "__main__":
    unittest.main()
