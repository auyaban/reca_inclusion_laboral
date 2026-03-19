import unittest

import text_review


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


if __name__ == "__main__":
    unittest.main()
