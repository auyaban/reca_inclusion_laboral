SECTION_2_5 = {
    "title": (
        "2.5 CONDICIONES DE ACCESIBILIDAD DISCAPACIDAD INTELECTUAL - TEA "
        "(TRASTORNO ESPECTRO AUTISTA)"
    ),
    "accesible_options": ["No", "Si", "Parcial"],
    "questions": [
        {
            "id": "material_seleccion_cognitiva",
            "label": "¿El material utilizado en el proceso de selección es accesible?",
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No aplica.",
            ],
        },
        {
            "id": "material_contratacion_cognitiva",
            "label": "¿El material utilizado en el proceso de contratación es accesible?",
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No aplica.",
            ],
        },
        {
            "id": "material_induccion_cognitiva",
            "label": "¿El material utilizado en el proceso de inducción y reinducción es accesible?",
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No aplica.",
            ],
        },
        {
            "id": "material_evaluacion_cognitiva",
            "label": "¿El material utilizado en el proceso de evaluación del desempeño es accesible?",
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No aplica.",
            ],
        },
        {
            "id": "ascensor_facil_ubicacion",
            "label": (
                "¿El ascensor al interior de las instalaciones es de fácil ubicación y su "
                "llamado a piso es de fácil entendimiento?"
            ),
            "type": "accesible_con_observaciones",
            "has_accesible": True,
        },
        {
            "id": "distribucion_zonas_comunes_percepcion",
            "label": (
                "¿La distribución de las zonas comunes (cafetería, oficinas, entre otras) "
                "permite una fácil orientación y percepción espacial?"
            ),
            "type": "accesible_con_observaciones",
            "has_accesible": True,
        },
        {
            "id": "plataformas_autogestion_intelectual",
            "label": "¿La organización cuenta con plataformas de autogestión?",
            "type": "lista_multiple",
            "has_accesible": True,
            "options": [
                "Trámites administrativos.",
                "Proceso de autocapacitación.",
                "Trámites administrativos y proceso de autocapacitación.",
                "No aplica.",
            ],
            "options_secondary": [
                "Cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No cuenta con accesibilidad cognitiva (lectura fácil - lenguaje sencillo).",
                "No aplica.",
            ],
        },
        {
            "id": "ajustes_razonables_intelectual",
            "label": (
                "¿La organización cuenta con la posibilidad de hacer ajustes razonables "
                "individualizados?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Cuenta con la posibilidad de flexibilizar rutinas laborales.",
                "No cuenta con la posibilidad de flexibilizar rutinas laborales.",
                "No aplica.",
            ],
            "text_observaciones": True,
        },
    ],
}

SECTION_2_6 = {
    "title": "2.6 CONDICIONES DE ACCESIBILIDAD DISCAPACIDAD PSICOSOCIAL",
    "accesible_options": ["No", "Si", "Parcial"],
    "questions": [
        {
            "id": "ajustes_razonables_psicosocial",
            "label": (
                "¿La organización cuenta con la posibilidad de hacer ajustes razonables "
                "individualizados?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Es posible flexibilizar los niveles de ruido que tenga en su puesto de trabajo.",
                "No es posible flexibilizar los niveles de ruido que tenga en su puesto de trabajo.",
                "No aplica.",
            ],
            "text_observaciones": True,
        },
    ],
}
