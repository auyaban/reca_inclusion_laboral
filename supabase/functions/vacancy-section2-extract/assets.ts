const SECTION_2_VACANCY_SEMANTIC_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["vacancy"],
  properties: {
    vacancy: {
      type: "object",
      additionalProperties: false,
      required: [
        "vacancy_name",
        "openings_count",
        "position_level",
        "gender_preference",
        "age_requirement_text",
        "work_modality_text",
        "work_location_text",
        "salary_text",
        "contract_signing_text",
        "tests_text",
        "contract_type",
        "additional_benefits_text",
        "gender_flexibility_text",
        "women_benefits_text",
        "certificate_requirement",
        "certificate_notes",
      ],
      properties: {
        vacancy_name: { type: ["string", "null"] },
        openings_count: { type: ["string", "null"] },
        position_level: { enum: ["administrative", "operational", "services", null] },
        gender_preference: {
          enum: ["male", "female", "male_female", "other", "indifferent", null],
        },
        age_requirement_text: { type: ["string", "null"] },
        work_modality_text: { type: ["string", "null"] },
        work_location_text: { type: ["string", "null"] },
        salary_text: { type: ["string", "null"] },
        contract_signing_text: { type: ["string", "null"] },
        tests_text: { type: ["string", "null"] },
        contract_type: {
          enum: [
            "fixed_term",
            "indefinite",
            "work_or_labor",
            "service_contract",
            "indefinite_presumptive",
            "appointment",
            "apprenticeship",
            "provisional_appointment",
            null,
          ],
        },
        additional_benefits_text: { type: ["string", "null"] },
        gender_flexibility_text: { type: ["string", "null"] },
        women_benefits_text: { type: ["string", "null"] },
        certificate_requirement: { enum: ["yes", "no", "in_process", null] },
        certificate_notes: { type: ["string", "null"] },
      },
    },
  },
} as const;

const SECTION_2_1_SCHEDULE_EXPERIENCE_SEMANTIC_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["schedule_experience"],
  properties: {
    schedule_experience: {
      type: "object",
      additionalProperties: false,
      required: [
        "schedule_type",
        "entry_time_text",
        "exit_time_text",
        "lunch_duration",
        "break_duration",
        "workdays_text",
        "flexible_days_text",
        "schedule_notes_text",
        "experience_requirement",
        "main_functions_text",
        "tools_and_equipment_text",
      ],
      properties: {
        schedule_type: { enum: ["fixed", "rotating", "flexible", null] },
        entry_time_text: { type: ["string", "null"] },
        exit_time_text: { type: ["string", "null"] },
        lunch_duration: { enum: ["15m", "30m", "45m", "1h", "2h", "not_applicable", null] },
        break_duration: { enum: ["15m", "30m", "45m", "1h", "not_applicable", null] },
        workdays_text: { type: ["string", "null"] },
        flexible_days_text: { type: ["string", "null"] },
        schedule_notes_text: { type: ["string", "null"] },
        experience_requirement: {
          enum: [
            "three_months",
            "six_months",
            "one_year",
            "one_and_half_years",
            "two_years",
            "two_and_half_years",
            "three_years",
            "four_years",
            "five_years",
            "internships_valid",
            "no_experience",
            "with_or_without_experience",
            null,
          ],
        },
        main_functions_text: { type: ["string", "null"] },
        tools_and_equipment_text: { type: ["string", "null"] },
      },
    },
  },
} as const;

const CANDIDATE_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: [
    "nombre_vacante",
    "numero_vacantes",
    "nivel_cargo",
    "genero",
    "edad",
    "modalidad_trabajo",
    "lugar_trabajo",
    "salario_asignado",
    "firma_contrato",
    "aplicacion_pruebas",
    "tipo_contrato",
    "beneficios_adicionales",
    "cargo_flexible_genero",
    "beneficios_mujeres",
    "requiere_certificado",
    "requiere_certificado_observaciones",
    "horarios_asignados",
    "hora_ingreso",
    "hora_salida",
    "tiempo_almuerzo",
    "break_descanso",
    "dias_laborables",
    "dias_flexibles",
    "observaciones",
    "experiencia_meses",
    "funciones_tareas",
    "herramientas_equipos",
  ],
  properties: {
    nombre_vacante: { type: ["string", "null"] },
    numero_vacantes: { type: ["string", "null"] },
    nivel_cargo: { type: ["string", "null"] },
    genero: { type: ["string", "null"] },
    edad: { type: ["string", "null"] },
    modalidad_trabajo: { type: ["string", "null"] },
    lugar_trabajo: { type: ["string", "null"] },
    salario_asignado: { type: ["string", "null"] },
    firma_contrato: { type: ["string", "null"] },
    aplicacion_pruebas: { type: ["string", "null"] },
    tipo_contrato: { type: ["string", "null"] },
    beneficios_adicionales: { type: ["string", "null"] },
    cargo_flexible_genero: { type: ["string", "null"] },
    beneficios_mujeres: { type: ["string", "null"] },
    requiere_certificado: { type: ["string", "null"] },
    requiere_certificado_observaciones: { type: ["string", "null"] },
    horarios_asignados: { type: ["string", "null"] },
    hora_ingreso: { type: ["string", "null"] },
    hora_salida: { type: ["string", "null"] },
    tiempo_almuerzo: { type: ["string", "null"] },
    break_descanso: { type: ["string", "null"] },
    dias_laborables: { type: ["string", "null"] },
    dias_flexibles: { type: ["string", "null"] },
    observaciones: { type: ["string", "null"] },
    experiencia_meses: { type: ["string", "null"] },
    funciones_tareas: { type: ["string", "null"] },
    herramientas_equipos: { type: ["string", "null"] },
  },
} as const;

export const SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: [
    "schema_version",
    "form_id",
    "section_id",
    "subsection_key",
    "audio_unit",
    "transcription_summary",
    "warnings",
    "semantic",
    "candidate",
  ],
  properties: {
    schema_version: { type: "integer" },
    form_id: { enum: ["condiciones_vacante"] },
    section_id: { enum: ["section_2", "section_2_1"] },
    subsection_key: { enum: ["section_2_vacancy", "section_2_1_schedule_experience"] },
    audio_unit: { enum: ["single_section"] },
    transcription_summary: { type: "string" },
    warnings: {
      type: "array",
      items: { type: "string" },
    },
    semantic: {
      type: "object",
      additionalProperties: false,
      required: ["section_2_vacancy", "section_2_1_schedule_experience"],
      properties: {
        section_2_vacancy: SECTION_2_VACANCY_SEMANTIC_SCHEMA,
        section_2_1_schedule_experience: SECTION_2_1_SCHEDULE_EXPERIENCE_SEMANTIC_SCHEMA,
      },
    },
    candidate: CANDIDATE_SCHEMA,
  },
} as const;

export const SUBSECTION_SPECS = {
  subsections: {
    section_2_vacancy: {
      title: "2. Caracteristicas de la vacante",
      script:
        "Responde en un solo audio y solo con informacion confirmada. Los campos abiertos deben quedar en texto fiel al audio; no fuerces formatos especiales para salario, fechas o edad.",
      questions: [
        "Como se llama la vacante y cuantas vacantes hay. Si esta claro, devuelve el numero de vacantes en digitos.",
        "Que nivel de cargo es, si hay preferencia de genero y que edad o rango de edad buscan.",
        "Cual es la modalidad, el lugar de trabajo y el salario.",
        "Cuando o como seria la firma del contrato, si aplican pruebas y que tipo de contrato es.",
        "Que beneficios adicionales ofrece la empresa.",
        "Si el cargo es flexible segun genero, si hay beneficios para mujeres y si requiere certificado de discapacidad con observacion si aplica.",
      ],
      examples: [
        "Vacante auxiliar logistico, son 2 vacantes, cargo operativo, genero indiferente, entre 20 y 45 anos, trabajo presencial en Fontibon, salario minimo con prestaciones, contratacion inmediata, si aplican pruebas psicotecnicas, contrato a termino fijo, beneficios de ruta y alimentacion, el cargo es flexible segun genero, hay beneficios para mujeres y el certificado de discapacidad esta en tramite.",
      ],
      prompt_fragment:
        "Usa semantic.section_2_vacancy.vacancy para los cuatro catalogos cerrados: position_level, gender_preference, contract_type y certificate_requirement. openings_count debe quedar en digitos cuando sea claro. Los campos abiertos deben quedarse en texto breve y fiel al audio. El salario puede venir en numeros o palabras. La firma del contrato puede venir como fecha exacta o expresion relativa como inmediata, la proxima semana o al cerrar proceso. Si un dato no es claro, dejalo en null.",
    },
    section_2_1_schedule_experience: {
      title: "2.1 Horarios, experiencia, funciones y herramientas",
      script:
        "Responde en un solo audio. Esta subseccion no cubre niveles educativos ni formacion academica. Solo cubre horarios, experiencia, funciones, herramientas y observaciones de jornada.",
      questions: [
        "Que horario asignado tiene la vacante: fijo, rotativo o con flexibilizacion.",
        "Cual es la hora de ingreso y la hora de salida.",
        "Cuanto dura el almuerzo y el break o descanso.",
        "Cuales son los dias laborables, si hay dias flexibles y si hay alguna observacion de la jornada.",
        "Cuanta experiencia pide la vacante.",
        "Cuales son las funciones principales y que herramientas o equipos se usan.",
      ],
      examples: [
        "Horarios rotativos, ingreso a las 6 de la manana y salida a las 2 de la tarde o 10 de la noche segun rotacion, almuerzo de 1 hora, break de 15 minutos, se trabaja de lunes a sabado con un dia compensatorio, viernes flexible cada 15 dias, experiencia de 1 ano, funciones operar maquinas, hacer control de calidad y reportar novedades, herramientas empacadoras, basculas y elementos de proteccion personal.",
      ],
      prompt_fragment:
        "Usa semantic.section_2_1_schedule_experience.schedule_experience para los catalogos cerrados: schedule_type, lunch_duration, break_duration y experience_requirement. Las horas, los dias, las funciones, herramientas y observaciones deben quedar en texto fiel al audio. No intentes llenar formacion academica ni niveles educativos en esta subseccion.",
    },
  },
} as const;

export const SYSTEM_PROMPT = `
Eres un extractor estructurado para el formulario Condiciones de Vacante.

Objetivo:
- Entender un audio corto en espanol y convertirlo a JSON estricto.
- No inventar datos.
- Mantener los campos abiertos en lenguaje natural breve y fiel al audio.

Reglas:
- Para section_2_vacancy usa semantic.section_2_vacancy.vacancy en los catalogos cerrados.
- Para section_2_1_schedule_experience usa semantic.section_2_1_schedule_experience.schedule_experience en los catalogos cerrados.
- En candidate puedes dejar texto libre para salario, edad, modalidad, lugar, horas, dias, funciones, herramientas y observaciones.
- No reformules de mas.
- No conviertas de manera obligatoria fechas, montos o rangos a formatos numericos, salvo numero_vacantes cuando sea claro.
- Si un dato no aparece o no es confiable, devuelve null.
`.trim();

export const CONTRACT_PROMPT = `
Subseccion section_2_vacancy:
- candidate: nombre_vacante, numero_vacantes, nivel_cargo, genero, edad, modalidad_trabajo, lugar_trabajo, salario_asignado, firma_contrato, aplicacion_pruebas, tipo_contrato, beneficios_adicionales, cargo_flexible_genero, beneficios_mujeres, requiere_certificado, requiere_certificado_observaciones.
- semantic.section_2_vacancy.vacancy: vacancy_name, openings_count, position_level, gender_preference, age_requirement_text, work_modality_text, work_location_text, salary_text, contract_signing_text, tests_text, contract_type, additional_benefits_text, gender_flexibility_text, women_benefits_text, certificate_requirement, certificate_notes.

Subseccion section_2_1_schedule_experience:
- candidate: horarios_asignados, hora_ingreso, hora_salida, tiempo_almuerzo, break_descanso, dias_laborables, dias_flexibles, observaciones, experiencia_meses, funciones_tareas, herramientas_equipos.
- semantic.section_2_1_schedule_experience.schedule_experience: schedule_type, entry_time_text, exit_time_text, lunch_duration, break_duration, workdays_text, flexible_days_text, schedule_notes_text, experience_requirement, main_functions_text, tools_and_equipment_text.
`.trim();
