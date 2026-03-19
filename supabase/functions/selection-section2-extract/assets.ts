export const CANDIDATE_FIELD_IDS = [
  "numero",
  "nombre_oferente",
  "cedula",
  "certificado_porcentaje",
  "discapacidad",
  "telefono_oferente",
  "resultado_certificado",
  "cargo_oferente",
  "nombre_contacto_emergencia",
  "parentesco",
  "telefono_emergencia",
  "fecha_nacimiento",
  "edad",
  "pendiente_otros_oferentes",
  "lugar_firma_contrato",
  "fecha_firma_contrato",
  "cuenta_pension",
  "tipo_pension",
  "desarrollo_actividad",
  "medicamentos_nivel_apoyo",
  "medicamentos_conocimiento",
  "medicamentos_horarios",
  "medicamentos_nota",
  "alergias_nivel_apoyo",
  "alergias_tipo",
  "alergias_nota",
  "restriccion_nivel_apoyo",
  "restriccion_conocimiento",
  "restriccion_nota",
  "controles_nivel_apoyo",
  "controles_asistencia",
  "controles_frecuencia",
  "controles_nota",
  "desplazamiento_nivel_apoyo",
  "desplazamiento_modo",
  "desplazamiento_transporte",
  "desplazamiento_nota",
  "ubicacion_nivel_apoyo",
  "ubicacion_ciudad",
  "ubicacion_aplicaciones",
  "ubicacion_nota",
  "dinero_nivel_apoyo",
  "dinero_reconocimiento",
  "dinero_manejo",
  "dinero_medios",
  "dinero_nota",
  "presentacion_nivel_apoyo",
  "presentacion_personal",
  "presentacion_nota",
  "comunicacion_escrita_nivel_apoyo",
  "comunicacion_escrita_apoyo",
  "comunicacion_escrita_nota",
  "comunicacion_verbal_nivel_apoyo",
  "comunicacion_verbal_apoyo",
  "comunicacion_verbal_nota",
  "decisiones_nivel_apoyo",
  "toma_decisiones",
  "toma_decisiones_nota",
  "aseo_nivel_apoyo",
  "alimentacion",
  "aseo_criar_apoyo",
  "aseo_comunicacion_apoyo",
  "aseo_ayudas_apoyo",
  "aseo_alimentacion",
  "aseo_movilidad_funcional",
  "aseo_higiene_aseo",
  "aseo_nota",
  "instrumentales_nivel_apoyo",
  "instrumentales_actividades",
  "instrumentales_criar_apoyo",
  "instrumentales_comunicacion_apoyo",
  "instrumentales_movilidad_apoyo",
  "instrumentales_finanzas",
  "instrumentales_cocina_limpieza",
  "instrumentales_crear_hogar",
  "instrumentales_salud_cuenta_apoyo",
  "instrumentales_nota",
  "actividades_nivel_apoyo",
  "actividades_apoyo",
  "actividades_esparcimiento_apoyo",
  "actividades_esparcimiento_cuenta_apoyo",
  "actividades_complementarios_apoyo",
  "actividades_complementarios_cuenta_apoyo",
  "actividades_subsidios_cuenta_apoyo",
  "actividades_nota",
  "discriminacion_nivel_apoyo",
  "discriminacion",
  "discriminacion_violencia_apoyo",
  "discriminacion_violencia_cuenta_apoyo",
  "discriminacion_vulneracion_apoyo",
  "discriminacion_vulneracion_cuenta_apoyo",
  "discriminacion_nota",
] as const;

const SECTION_2_IDENTITY_SCHEMA = {
  type: ["object", "null"],
  additionalProperties: false,
  required: ["identity"],
  properties: {
    identity: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["document_number", "applicant_phone", "emergency_phone", "birthdate_iso"],
      properties: {
        document_number: { type: ["string", "null"] },
        applicant_phone: { type: ["string", "null"] },
        emergency_phone: { type: ["string", "null"] },
        birthdate_iso: { type: ["string", "null"] },
      },
    },
  },
} as const;

const SECTION_4_1_SUPPORT_LEVELS = ["none", "low", "medium", "high", "not applicable", null];

const SECTION_4_1_HEALTH_SCHEMA = {
  type: ["object", "null"],
  additionalProperties: false,
  required: ["medications", "allergies", "restrictions", "specialist_controls"],
  properties: {
    medications: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "status", "schedule_status", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        status: {
          enum: ["not taking", "self managed", "third party managed", "unknown", "not applicable", null],
        },
        schedule_status: {
          enum: ["self managed", "third party managed", "unknown", "not applicable", null],
        },
        details: { type: ["string", "null"] },
      },
    },
    allergies: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "status", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        status: {
          enum: ["none reported", "self managed", "unknown", "described", "not applicable", null],
        },
        details: { type: ["string", "null"] },
      },
    },
    restrictions: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "status", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        status: {
          enum: [
            "none reported",
            "self managed",
            "unknown",
            "does not know management",
            "not applicable",
            null,
          ],
        },
        details: { type: ["string", "null"] },
      },
    },
    specialist_controls: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "attendance", "frequency", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        attendance: {
          enum: ["attends and self manages", "attends", "unknown", "not applicable", null],
        },
        frequency: {
          enum: ["monthly", "quarterly", "semiannual", "other", "not applicable", null],
        },
        details: { type: ["string", "null"] },
      },
    },
  },
} as const;

const SECTION_4_2_A_SKILLS_SCHEMA = {
  type: ["object", "null"],
  additionalProperties: false,
  required: [
    "mobility",
    "orientation",
    "money",
    "presentation",
    "written_communication",
    "verbal_communication",
    "decision_making",
  ],
  properties: {
    mobility: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "mode", "transport", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        mode: {
          enum: ["autonomous", "temporary support", "permanent support", "third party support", "not applicable", null],
        },
        transport: {
          enum: ["walking", "bicycle", "mass transit", "own vehicle", "special vehicle", "not applicable", null],
        },
        details: { type: ["string", "null"] },
      },
    },
    orientation: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "city_status", "references_status", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        city_status: { enum: ["autonomous", "apps", "accompanied", "does not orient", null] },
        references_status: { enum: ["references", "no references", "cardinal points", "not applicable", null] },
        details: { type: ["string", "null"] },
      },
    },
    money: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "recognition", "management", "mediums", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        recognition: { enum: ["autonomous", "family support", null] },
        management: {
          enum: ["autonomous", "occasional support", "recognizes only", "does not recognize", "not applicable", null],
        },
        mediums: {
          type: ["array", "null"],
          items: { enum: ["cash", "card", "digital"] },
        },
        details: { type: ["string", "null"] },
      },
    },
    presentation: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "dress_code", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        dress_code: {
          enum: ["appropriate", "appropriate with improvement", "partially appropriate", "inappropriate", "not applicable", null],
        },
        details: { type: ["string", "null"] },
      },
    },
    written_communication: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "support_status", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        support_status: { enum: ["knows and uses", "uses some", "knows not uses", "neither", "not applicable", null] },
        details: { type: ["string", "null"] },
      },
    },
    verbal_communication: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "support_status", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        support_status: { enum: ["knows and uses", "uses some", "knows not uses", "neither", "not applicable", null] },
        details: { type: ["string", "null"] },
      },
    },
    decision_making: {
      type: ["object", "null"],
      additionalProperties: false,
      required: ["support_level", "decision_status", "details"],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        decision_status: {
          enum: ["autonomous", "occasional support", "consults third party", "requires third party", "not applicable", null],
        },
        details: { type: ["string", "null"] },
      },
    },
  },
} as const;

const SEMANTIC_BINARY_ENUM = ["yes", "no", "not applicable", null] as const;

const SECTION_4_2_B_SUPPORT_SCHEMA = {
  type: ["object", "null"],
  additionalProperties: false,
  required: ["daily_living", "instrumental", "work_activities", "discrimination"],
  properties: {
    daily_living: {
      type: ["object", "null"],
      additionalProperties: false,
      required: [
        "support_level",
        "scope",
        "child_care",
        "communication_systems",
        "assistive_devices",
        "feeding",
        "functional_mobility",
        "hygiene",
        "details",
      ],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        scope: { enum: ["none", "some", "most", "all", "not applicable", null] },
        child_care: { enum: SEMANTIC_BINARY_ENUM },
        communication_systems: { enum: SEMANTIC_BINARY_ENUM },
        assistive_devices: { enum: SEMANTIC_BINARY_ENUM },
        feeding: { enum: SEMANTIC_BINARY_ENUM },
        functional_mobility: { enum: SEMANTIC_BINARY_ENUM },
        hygiene: { enum: SEMANTIC_BINARY_ENUM },
        details: { type: ["string", "null"] },
      },
    },
    instrumental: {
      type: ["object", "null"],
      additionalProperties: false,
      required: [
        "support_level",
        "scope",
        "child_care",
        "communication_systems",
        "community_mobility",
        "finances",
        "cooking_cleaning",
        "household",
        "health_support",
        "details",
      ],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        scope: { enum: ["none", "some", "most", "all", "not applicable", null] },
        child_care: { enum: SEMANTIC_BINARY_ENUM },
        communication_systems: { enum: SEMANTIC_BINARY_ENUM },
        community_mobility: { enum: SEMANTIC_BINARY_ENUM },
        finances: { enum: SEMANTIC_BINARY_ENUM },
        cooking_cleaning: { enum: SEMANTIC_BINARY_ENUM },
        household: { enum: SEMANTIC_BINARY_ENUM },
        health_support: { enum: SEMANTIC_BINARY_ENUM },
        details: { type: ["string", "null"] },
      },
    },
    work_activities: {
      type: ["object", "null"],
      additionalProperties: false,
      required: [
        "support_level",
        "scope",
        "family_recreation_requires_support",
        "family_recreation_has_support",
        "medical_followup_requires_support",
        "medical_followup_has_support",
        "children_subsidies_has_support",
        "details",
      ],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        scope: { enum: ["none", "some", "most", "all", "not applicable", null] },
        family_recreation_requires_support: { enum: SEMANTIC_BINARY_ENUM },
        family_recreation_has_support: { enum: SEMANTIC_BINARY_ENUM },
        medical_followup_requires_support: { enum: SEMANTIC_BINARY_ENUM },
        medical_followup_has_support: { enum: SEMANTIC_BINARY_ENUM },
        children_subsidies_has_support: { enum: SEMANTIC_BINARY_ENUM },
        details: { type: ["string", "null"] },
      },
    },
    discrimination: {
      type: ["object", "null"],
      additionalProperties: false,
      required: [
        "support_level",
        "scope",
        "physical_violence_requires_support",
        "physical_violence_has_support",
        "rights_violation_requires_support",
        "rights_violation_has_support",
        "details",
      ],
      properties: {
        support_level: { enum: SECTION_4_1_SUPPORT_LEVELS },
        scope: { enum: ["none", "some contexts", "repeated", "lifelong", "not applicable", null] },
        physical_violence_requires_support: { enum: SEMANTIC_BINARY_ENUM },
        physical_violence_has_support: { enum: SEMANTIC_BINARY_ENUM },
        rights_violation_requires_support: { enum: SEMANTIC_BINARY_ENUM },
        rights_violation_has_support: { enum: SEMANTIC_BINARY_ENUM },
        details: { type: ["string", "null"] },
      },
    },
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
    schema_version: { type: "integer", const: 1 },
    form_id: { type: "string", const: "seleccion_incluyente_labs" },
    section_id: { type: "string", const: "section_2" },
    subsection_key: {
      type: "string",
      enum: [
        "section_2_fields",
        "section_3_desarrollo",
        "section_4_1_salud",
        "section_4_2_a_habilidades",
        "section_4_2_b_actividades",
      ],
    },
    audio_unit: { type: "string", const: "single_candidate" },
    transcription_summary: { type: "string" },
    warnings: {
      type: "array",
      items: { type: "string" },
    },
    semantic: {
      type: "object",
      additionalProperties: false,
      required: ["section_2_identity", "section_4_1_health", "section_4_2_a_skills", "section_4_2_b_support"],
      properties: {
        section_2_identity: SECTION_2_IDENTITY_SCHEMA,
        section_4_1_health: SECTION_4_1_HEALTH_SCHEMA,
        section_4_2_a_skills: SECTION_4_2_A_SKILLS_SCHEMA,
        section_4_2_b_support: SECTION_4_2_B_SUPPORT_SCHEMA,
      },
    },
    candidate: {
      type: "object",
      additionalProperties: false,
      required: [...CANDIDATE_FIELD_IDS],
      properties: Object.fromEntries(
        CANDIDATE_FIELD_IDS.map((fieldId) => [fieldId, { type: ["string", "null"] }]),
      ),
    },
  },
};

export const SYSTEM_PROMPT = `Eres un extractor de datos para el formulario seleccion_incluyente_labs, seccion section_2.

Devuelve solo JSON valido del schema.

Reglas:
- Usa null cuando falte informacion.
- No inventes datos.
- Usa exactamente las opciones validas de los dropdowns cuando decidas llenar candidate y el audio sea claro.
- numero y edad siempre van en null.
- En section_2_fields, semantic.section_2_identity es la capa principal para cedula, telefonos y fecha de nacimiento. Usa solo digitos para cedula y telefonos; si un numero supera 10 digitos o no es confiable, dejalo en null. birthdate_iso debe ir en YYYY-MM-DD cuando sea claro.
- En section_4_1_salud, semantic.section_4_1_health es la capa principal de interpretacion. Usa solo los estados semanticos definidos alli; si el dropdown exacto de candidate no es obvio, dejalo en null.
- En section_4_2_a_habilidades, semantic.section_4_2_a_skills es la capa principal de interpretacion. Usa estados semanticos simples y deja candidate en null si el dropdown exacto no es obvio.
- En section_4_2_b_actividades, semantic.section_4_2_b_support es la capa principal de interpretacion. Usa estados semanticos simples para el alcance del apoyo y los subitems binarios.
- En 4.2A y 4.2B, si el audio deja clara la opcion principal numerada, no hace falta repetir el nivel general.
- Si el profesional dice solo, unicamente, ninguna, ninguno, lo demas no o una variante equivalente, puedes propagar No dentro del subgrupo correspondiente.
- Si enumera algunos positivos sin cerrar explicitamente el resto, los no mencionados quedan en null.
- Conserva detalles utiles en campos *_nota.
- warnings debe incluir ambiguedades, contradicciones o mezcla de varias personas.`;

export const CONTRACT_PROMPT = `Mapa corto del formulario:
- section_2_fields: nombre_oferente, cedula, certificado_porcentaje, discapacidad, telefono_oferente, resultado_certificado, cargo_oferente, nombre_contacto_emergencia, parentesco, telefono_emergencia, fecha_nacimiento, pendiente_otros_oferentes, lugar_firma_contrato, fecha_firma_contrato, cuenta_pension, tipo_pension.
- desarrollo: solo desarrollo_actividad.
- salud: medicamentos, alergias, restricciones y controles, con notas breves.
- habilidades 4.2A: desplazamiento, ubicacion, dinero, presentacion, comunicacion escrita, comunicacion verbal y toma de decisiones.
- actividades 4.2B: vida diaria, instrumentales, laborales y discriminacion.

Valores clave de section_2_fields:
- resultado_certificado: Aprobado | No aprobado | Pendiente
- pendiente_otros_oferentes: Si | No | Por Confirmar
- cuenta_pension: Si | No | Por Confirmar
- tipo_pension: Pension Invalidez | Subsidiada | Especial de vejez | Victimas conflicto | Familiar | Regimen especial | No aplica
- discapacidad: Discapacidad visual pérdida total de la visión | Discapacidad visual baja visión | Discapacidad auditiva | Discapacidad auditiva hipoacusia | Trastorno de espectro autista | Discapacidad intelectual | Discapacidad física | Discapacidad física usuario en silla de ruedas | Discapacidad psicosocial | Discapacidad múltiple | No aplica

Valores clave de salud y habilidades:
- niveles de apoyo: 0. No requiere apoyo. | 1. Nivel de apoyo Bajo. | 2. Nivel de apoyo medio. | 3. Nivel de apoyo alto. | No aplica.
- subitems binarios: Si | No | No aplica
- frecuencias de controles: Mensual | Trimestral | Semestral | Otra frecuencia | No aplica

Reglas de interpretacion:
- no sabe -> usa la opcion de desconocimiento si existe; si no existe, deja null y usa la nota
- no aplica -> solo cuando el profesional lo diga claramente
- fecha_firma_contrato puede ser una fecha real o el texto Por Confirmar si el profesional dice pendiente, por confirmar o equivalente
- en section_2_fields llena semantic.section_2_identity aunque candidate quede mayormente en null para cedula, telefonos y fecha_nacimiento
- semantic.section_2_identity.identity.document_number: string solo con digitos y maximo 10 caracteres
- semantic.section_2_identity.identity.applicant_phone: string solo con digitos y exactamente 10 caracteres cuando el telefono sea claro
- semantic.section_2_identity.identity.emergency_phone: string solo con digitos y exactamente 10 caracteres cuando el telefono sea claro
- semantic.section_2_identity.identity.birthdate_iso: fecha en formato YYYY-MM-DD
- en section_4_1_salud llena semantic.section_4_1_health aunque candidate quede mayormente en null
- semantic.section_4_1_health.medications.support_level: none | low | medium | high | not applicable
- semantic.section_4_1_health.medications.status: not taking | self managed | third party managed | unknown | not applicable
- semantic.section_4_1_health.medications.schedule_status: self managed | third party managed | unknown | not applicable
- semantic.section_4_1_health.allergies.status: none reported | self managed | unknown | described | not applicable
- semantic.section_4_1_health.restrictions.status: none reported | self managed | unknown | does not know management | not applicable
- semantic.section_4_1_health.specialist_controls.attendance: attends and self manages | attends | unknown | not applicable
- semantic.section_4_1_health.specialist_controls.frequency: monthly | quarterly | semiannual | other | not applicable
- en section_4_2_a_habilidades llena semantic.section_4_2_a_skills aunque candidate quede mayormente en null
- semantic.section_4_2_a_skills.mobility.mode: autonomous | temporary support | permanent support | third party support | not applicable
- semantic.section_4_2_a_skills.mobility.transport: walking | bicycle | mass transit | own vehicle | special vehicle | not applicable
- semantic.section_4_2_a_skills.orientation.city_status: autonomous | apps | accompanied | does not orient
- semantic.section_4_2_a_skills.orientation.references_status: references | no references | cardinal points | not applicable
- semantic.section_4_2_a_skills.money.recognition: autonomous | family support
- semantic.section_4_2_a_skills.money.management: autonomous | occasional support | recognizes only | does not recognize | not applicable
- semantic.section_4_2_a_skills.money.mediums: lista con cash | card | digital
- semantic.section_4_2_a_skills.presentation.dress_code: appropriate | appropriate with improvement | partially appropriate | inappropriate | not applicable
- semantic.section_4_2_a_skills.written_communication.support_status y verbal_communication.support_status: knows and uses | uses some | knows not uses | neither | not applicable
- semantic.section_4_2_a_skills.decision_making.decision_status: autonomous | occasional support | consults third party | requires third party | not applicable
- en section_4_2_b_actividades llena semantic.section_4_2_b_support aunque candidate quede mayormente en null
- semantic.section_4_2_b_support.daily_living.scope, instrumental.scope y work_activities.scope: none | some | most | all | not applicable
- semantic.section_4_2_b_support.discrimination.scope: none | some contexts | repeated | lifelong | not applicable
- los subitems binarios de section_4_2_b_support usan: yes | no | not applicable
- en 4.2A y 4.2B el desktop puede derivar nivel_apoyo desde el dropdown principal numerado cuando exista
- si un grupo principal de 4.2B queda en 0 o No aplica, el desktop puede completar subitems faltantes como No o No aplica
- no conviertas texto narrativo en negativos no mencionados`;

export const SUBSECTION_SPECS = {
  subsections: {
    section_2_fields: {
      title: "2. Datos del oferente",
      script:
        "Nombre, cedula, certificado, discapacidad, telefono, resultado, cargo, contacto de emergencia, parentesco, telefono, fecha de nacimiento, pendiente otros oferentes, pension y firma de contrato.",
      questions: [
        "Cual es el nombre completo del oferente y su cedula?",
        "Cual es el porcentaje del certificado y el tipo de discapacidad?",
        "Cual es el telefono del oferente y el cargo o vacante?",
        "Quien es el contacto de emergencia, que parentesco tiene y cual es su telefono?",
        "Cual es la fecha de nacimiento?",
        "Esta pendiente con otros oferentes, tiene pension y ya se confirmo la firma de contrato?",
      ],
      examples: [
        "Juan Perez, cedula 10203040, certificado 45 por ciento, discapacidad fisica, telefono 3001234567, resultado aprobado, cargo auxiliar logistico, contacto de emergencia Ana Perez madre 3105556677, fecha de nacimiento 12 de mayo de 1998, pendiente otros oferentes por confirmar, cuenta con pension no, firma de contrato en Bogota el 20 de marzo de 2026.",
      ],
      prompt_fragment:
        "Prioriza semantic.section_2_identity para cedula, telefonos y fecha_nacimiento. Usa solo digitos en cedula y telefonos. Si cedula o un telefono supera 10 digitos, queda en null. Para fecha_nacimiento usa semantic.section_2_identity.identity.birthdate_iso en YYYY-MM-DD cuando sea clara; el desktop derivara edad desde esa fecha. Si cuenta_pension es Por Confirmar, no inventes tipo_pension. fecha_firma_contrato puede ser una fecha o Por Confirmar; no inventes una fecha dura.",
    },
    section_3_desarrollo: {
      title: "3. Desarrollo de la actividad",
      script:
        "Di un relato corto y corrido: que se hizo, que se observo, que acuerdos quedaron y que pendientes hay. No hace falta enumerar campos.",
      questions: [
        "Que se hizo durante la actividad o entrevista?",
        "Que se observo del oferente?",
        "Que acuerdos o conclusiones quedaron?",
        "Que pendientes siguen abiertos?",
      ],
      examples: [
        "Se realizo entrevista individual, se reviso experiencia previa y se explico el proceso de seleccion. Se observo buena disposicion y quedaron pendientes documentos por confirmar.",
      ],
      prompt_fragment:
        "Extrae un unico texto claro en desarrollo_actividad. Conserva el orden narrativo del profesional. Limpia solo errores menores de organizacion del texto; no resumas, no agregues conclusiones y no conviertas este bloque en listas de dropdowns.",
    },
    section_4_1_salud: {
      title: "4.1 Condiciones medicas y de salud",
      script:
        "Habla por bloques: medicamentos, alergias, restricciones y controles. En cada bloque empieza con una salida corta como no requiere apoyo, no aplica, no sabe o si, y luego da el detalle breve si hace falta.",
      questions: [
        "Toma medicamentos? Si si, cuales, quien los conoce y conoce los horarios?",
        "Tiene alergias? Si si, cuales y sabe manejarlas?",
        "Tiene restricciones medicas o laborales? Si si, cuales y sabe manejarlas?",
        "Asiste a controles con especialista? Si si, cada cuanto y sabe como manejarlos?",
      ],
      examples: [
        "Medicamentos: no requiere apoyo, conoce los medicamentos y conoce los horarios, toma losartan. Alergias: no presenta alergias. Restricciones: nivel bajo, si tiene restricciones medicas y conoce su manejo, no cargar peso. Controles: nivel bajo, asiste a controles medicos con especialista y conoce el manejo, frecuencia trimestral.",
      ],
      prompt_fragment:
        "Prioriza semantic.section_4_1_health. Usa los estados semanticos simples de salud para medicamentos, alergias, restricciones y controles, y deja candidate en null cuando el dropdown exacto no sea completamente obvio. Si el profesional dice no sabe, usa el estado de desconocimiento correspondiente. Si un bloque es no aplica, no inventes detalles extra. Guarda detalles clinicos breves en semantic.details y en *_nota cuando sea util.",
    },
    section_4_2_a_habilidades: {
      title: "4.2A Habilidades basicas de la vida diaria",
      script:
        "Habla bloque por bloque: desplazamiento, ubicacion, dinero, presentacion personal, comunicacion escrita, comunicacion verbal y toma de decisiones. No hace falta dictar el nivel general si la situacion principal ya queda clara.",
      questions: [
        "Como se desplaza la persona y que transporte usa?",
        "Como se ubica en la ciudad? Usa aplicaciones, referencias o necesita acompanamiento?",
        "Reconoce y maneja el dinero? Que medios usa: fisico, tarjeta o digital?",
        "Su presentacion personal es acorde al contexto o requiere mejora?",
        "Conoce y maneja apoyos de comunicacion escrita?",
        "Conoce y maneja apoyos de comunicacion verbal?",
        "Toma decisiones sola o consulta a un tercero?",
      ],
      examples: [
        "Desplazamiento: se desplaza de forma independiente con apoyo temporal, usa baston y Transmilenio. Ubicacion: se ubica usando Maps y direcciones. Dinero: reconoce el dinero con apoyo familiar, lo maneja con apoyo ocasional y usa dinero fisico y plastico. Presentacion personal: acorde al contexto. Comunicacion escrita: conoce pero no maneja apoyos. Comunicacion verbal: no requiere apoyo. Toma decisiones: en ocasiones consulta a la madre.",
      ],
      prompt_fragment:
        "Prioriza semantic.section_4_2_a_skills. Usa estados semanticos simples para desplazamiento, ubicacion, dinero, presentacion, comunicacion y toma de decisiones, y deja candidate en null cuando el dropdown exacto no sea completamente obvio. El desktop puede derivar nivel_apoyo localmente. Usa semantic.details y *_nota para detalles como baston, apps o a quien consulta.",
    },
    section_4_2_b_actividades: {
      title: "4.2B Actividades, apoyos y discriminacion",
      script:
        "Habla grupo por grupo: vida diaria, instrumentales, actividades laborales y discriminacion. Si todo el grupo es negativo basta con decir no requiere apoyo o no ha sufrido discriminacion. Si solo algunas opciones van en Si, di cuales son y cierra con lo demas no.",
      questions: [
        "En actividades de la vida diaria requiere apoyo? Si si, en cuales: crianza, comunicacion, ayudas tecnicas, alimentacion, movilidad o higiene?",
        "En actividades instrumentales requiere apoyo? Si si, en cuales: crianza, comunicacion, movilidad, finanzas, cocina, hogar o salud?",
        "En actividades laborales requiere apoyo? Si si, en cuales y cuenta con apoyo?",
        "Ha sufrido discriminacion? Si si, en que nivel y si hubo violencia fisica o vulneracion de derechos, requiere apoyo o cuenta con apoyo?",
      ],
      examples: [
        "Vida diaria: requiere apoyo en algunas actividades, solo alimentacion; lo demas no. Instrumentales: si necesita apoyo en finanzas y cocina; lo demas no. Actividades laborales: complementarios medicos si, lo demas no. Discriminacion: ha sufrido discriminacion en algunos contextos, violencia fisica no, vulneracion de derechos no.",
      ],
      prompt_fragment:
        "Prioriza semantic.section_4_2_b_support. Usa estados semanticos simples para el alcance del apoyo del grupo y para cada subitem binario. Si el grupo principal queda en 0 o No aplica, el desktop puede completar subitems faltantes como No o No aplica. Si el profesional dice solo X o lo demas no, puedes marcar el resto en No. Si no lo dice, deja los no mencionados en null. Nunca inventes negativos.",
    },
  },
};
