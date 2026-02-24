SECTION_3 = {
    "title": "3. CONDICIONES ORGANIZACIONALES",
    "accesible_options": ["No", "Si", "Parcial"],
    "questions": [
        {
            "id": "experiencia_vinculacion_pcd",
            "label": "¿La empresa ha contratado o tiene experiencias de vinculación PcD?",
            "type": "accesible_con_observaciones",
            "has_accesible": True,
        },
        {
            "id": "personal_tercerizado_capacitado",
            "label": (
                "¿El personal tercerizado (Seguridad, servicios generales, mantenimiento, "
                "entre otros) está capacitado o tiene experiencia en interacción con PcD?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Cuenta con capacitación.",
                "Cuenta con experiencia.",
                "Cuenta con capacitación y experiencia.",
                "No aplica.",
            ],
        },
        {
            "id": "personal_directo_capacitado",
            "label": (
                "¿El personal directo de la empresa está capacitado o tiene experiencia en "
                "interacción con PcD?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Cuenta con capacitación.",
                "Cuenta con experiencia.",
                "Cuenta con capacitación y experiencia.",
                "No aplica.",
            ],
        },
        {
            "id": "apoyo_arl_seguridad",
            "label": (
                "¿La empresa ha solicitado el apoyo y servicios de la ARL para velar por la "
                "seguridad y bienestar de los trabajadores en condición de discapacidad?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "La ARL apoya mensual.",
                "La ARL apoya Trimestral.",
                "La ARL apoya Semestral.",
                "La ARL apoya Anual.",
                "Nunca lo han solicitado.",
                "No aplica.",
            ],
        },
        {
            "id": "capacitacion_emergencias",
            "label": (
                "¿La empresa ha sido capacitada en plan de emergencia y evacuación, ante una "
                "emergencia por bomberos?"
            ),
            "type": "accesible_con_observaciones",
            "has_accesible": True,
        },
        {
            "id": "politica_diversidad_inclusion",
            "label": "¿La empresa cuenta con una política de diversidad e inclusión laboral?",
            "type": "lista",
            "has_accesible": True,
            "options": [
                "La política se encuentra finalizada.",
                "La política está en construcción.",
                "No aplica.",
            ],
        },
        {
            "id": "rrhh_normatividad",
            "label": (
                "¿El área de recursos humanos conoce la normatividad vigente para la "
                "vinculación laboral de PcD?"
            ),
            "type": "lista_multiple",
            "has_accesible": True,
            "options": [
                "Conoce los beneficios tributarios.",
                "No conoce los beneficios tributarios.",
                "No aplica.",
            ],
            "options_secondary": [
                "Conoce los beneficios en licitaciones públicas.",
                "No conoce los beneficios en licitaciones públicas.",
                "No aplica.",
            ],
            "options_tertiary": [
                "Conoce los beneficios en cuota de aprendiz SENA.",
                "No conoce los beneficios en cuota de aprendiz SENA.",
                "No aplica.",
            ],
            "options_quaternary": [
                "Conoce la normatividad referente al certificado de discapacidad.",
                "No conoce la normatividad referente al certificado de discapacidad.",
                "No aplica.",
            ],
        },
        {
            "id": "ajustes_razonables_empresa",
            "label": (
                "¿La empresa tiene dentro de su alcance realizar ajustes razonables, e "
                "implementar sistemas de apoyo que se sugieren para la vinculación de PcD?"
            ),
            "type": "lista_multiple",
            "has_accesible": True,
            "accesible_options": ["Parcial", "Si", "No"],
            "options": [
                "Es posible realizar ajustes al puesto de trabajo.",
                "No es posible realizar ajustes al puesto de trabajo.",
                "No aplica.",
            ],
            "options_secondary": [
                "Es posible realizar ajustes arquitectónicos.",
                "No es posible realizar ajustes arquitectónicos.",
                "No aplica.",
            ],
            "options_tertiary": [
                "Es posible realizar ajustes en documentación y presentación de la información.",
                "No es posible realizar ajustes en documentación y presentación de la información.",
                "No aplica.",
            ],
            "options_quaternary": [
                "Es posible realizar ajustes en cuanto a carga laboral o trabajo bajo presión.",
                "No es posible realizar ajustes en cuanto a carga laboral o trabajo bajo presión.",
                "No aplica.",
            ],
        },
        {
            "id": "protocolo_emergencias_pcd",
            "label": (
                "¿La empresa cuenta con protocolo de atención de emergencias para personas "
                "con discapacidad?"
            ),
            "type": "lista_multiple",
            "has_accesible": True,
            "accesible_options": ["Parcial", "Si", "No"],
            "options": [
                "Cuenta con los números de contacto en caso de emergencia.",
                "No cuenta con los números de contacto en caso de emergencia.",
                "No aplica.",
            ],
            "options_secondary": [
                "Conoce los puntos de atención en caso de emergencia (clinicas, hospitales).",
                "No conoce los puntos de atención en caso de emergencia (clinicas, hospitales).",
                "No aplica.",
            ],
            "options_tertiary": [
                "Tiene conocimiento sobre las recomendaciones médicas asociadas a la PcD.",
                "No tiene conocimiento sobre las recomendaciones médicas asociadas a la PcD.",
                "No aplica.",
            ],
        },
        {
            "id": "apoyo_bomberos_discapacidad",
            "label": (
                "¿La empresa ha solicitado apoyo de otra entidad como bomberos para población "
                "con discapacidad?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Si ha contado con apoyo.",
                "No ha contado con apoyo.",
                "Se cuenta con apoyo parcial.",
                "No aplica.",
            ],
            "text_observaciones": True,
        },
        {
            "id": "disponibilidad_tiempo_inclusion",
            "label": (
                "¿La empresa cuenta con disposición y tiempo para realizar los procesos de "
                "Inclusión Laboral?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Se cuenta con 30 minutos.",
                "Se cuenta con 45 minutos.",
                "Se cuenta con 60 minutos o más.",
            ],
        },
        {
            "id": "practicas_equidad_genero",
            "label": (
                "¿La empresa cuenta con prácticas inclusivas orientadas a la equidad de género "
                "y cuáles?"
            ),
            "type": "lista",
            "has_accesible": True,
            "options": [
                "Ajuste de lenguaje inclusivo en las comunicaciones internas.",
                "Salas de lactancia.",
                "Guarderías.",
                "Programas de apoyo a mujeres cuidadoras.",
                "Otros.",
            ],
        },
    ],
}
