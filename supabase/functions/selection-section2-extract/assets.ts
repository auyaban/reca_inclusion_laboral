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
- Usa exactamente las opciones validas de los dropdowns cuando el audio sea claro.
- numero y edad siempre van en null.
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
- no conviertas texto narrativo en negativos no mencionados`;

export const SUBSECTION_SPECS = {
  subsections: {
    section_2_fields: {
      title: "2. Datos del oferente",
      script:
        "Nombre, cedula, certificado, discapacidad, telefono, resultado, cargo, contacto de emergencia, parentesco, telefono, fecha de nacimiento, pendiente otros oferentes, pension y firma de contrato.",
      examples: [
        "Juan Perez, cedula 10203040, certificado 45 por ciento, discapacidad fisica, telefono 3001234567, resultado aprobado, cargo auxiliar logistico, contacto de emergencia Ana Perez madre 3105556677, fecha de nacimiento 12 de mayo de 1998, pendiente otros oferentes no, cuenta con pension no, firma de contrato en Bogota el 20 de marzo de 2026.",
        "Maria Lopez, cedula 51880022, certificado 33 por ciento, trastorno de espectro autista, telefono 3019876543, resultado pendiente, cargo operaria de empaque, contacto de emergencia Luis Lopez hermano 3124442233, fecha de nacimiento 3 de agosto de 2001, pendiente otros oferentes por confirmar, cuenta con pension por confirmar.",
        "Laura Diaz, cedula 50660011, discapacidad fisica, telefono 3005551122, resultado aprobado, cargo auxiliar administrativa, contacto de emergencia Marta Diaz esposa 3104448899, fecha de nacimiento 22 de noviembre de 1996, pendiente otros oferentes no, firma de contrato pendiente por confirmar, cuenta con pension no.",
      ],
      prompt_fragment:
        "Extrae solo datos cortos del oferente. numero y edad quedan en null. Si cuenta_pension es Por Confirmar, no inventes tipo_pension. fecha_firma_contrato puede ser una fecha o Por Confirmar; no inventes una fecha dura.",
    },
    section_3_desarrollo: {
      title: "3. Desarrollo de la actividad",
      script: "Describe que se hizo, que se observo y que pendientes quedaron.",
      examples: [
        "Se realizo entrevista individual, se reviso experiencia previa y se explico el proceso de seleccion.",
        "Se valido hoja de vida, motivacion frente al cargo y apoyos iniciales requeridos.",
      ],
      prompt_fragment:
        "Extrae un unico texto claro en desarrollo_actividad. No conviertas este bloque en listas de dropdowns.",
    },
    section_4_1_salud: {
      title: "4.1 Condiciones medicas y de salud",
      script: "Medicamentos, alergias, restricciones y controles. Usa frases cortas como no requiere apoyo, no sabe o no aplica.",
      examples: [
        "Medicamentos: no requiere apoyo, conoce los medicamentos y conoce los horarios, toma losartan. Alergias: no presenta alergias. Restricciones: nivel bajo, si tiene restricciones medicas y conoce su manejo, no cargar peso. Controles: nivel bajo, asiste a controles medicos con especialista y conoce el manejo, frecuencia trimestral.",
        "Medicamentos: no aplica. Alergias: nivel medio, presenta alergias y sabe darle manejo, alergia a penicilina. Restricciones: no sabe si tiene restricciones medicas. Controles: no sabe si tiene controles medicos con especialista.",
      ],
      prompt_fragment:
        "Usa las opciones exactas de medicamentos, alergias, restricciones y controles. Guarda detalles clinicos breves en *_nota.",
    },
    section_4_2_a_habilidades: {
      title: "4.2A Habilidades basicas de la vida diaria",
      script:
        "Desplazamiento, ubicacion, dinero, presentacion personal, comunicacion escrita, comunicacion verbal y toma de decisiones.",
      examples: [
        "Se desplaza de forma independiente con apoyo temporal, usa baston y Transmilenio. Se ubica usando Maps. Maneja dinero con apoyo ocasional y usa dinero fisico y plastico. Presentacion personal acorde al contexto. Comunicacion escrita: conoce pero no maneja apoyos. Comunicacion verbal: no requiere apoyo. Toma decisiones pero a veces consulta a la madre.",
        "No se desplaza de manera independiente, requiere acompanamiento de un tercero y vehiculo especial. Requiere acompanamiento para ubicarse. Solo reconoce el dinero pero no lo sabe manejar. No conoce ni maneja apoyos de comunicacion verbal ni escrita.",
      ],
      prompt_fragment:
        "Extrae solo desplazamiento a toma de decisiones. Usa la nota para detalles como baston, apps o a quien consulta.",
    },
    section_4_2_b_actividades: {
      title: "4.2B Actividades, apoyos y discriminacion",
      script:
        "Vida diaria, instrumentales, laborales y discriminacion. Si solo algunas son Si, el profesional debe decir lo demas no.",
      examples: [
        "Vida diaria: requiere apoyo en algunas actividades, solo alimentacion; lo demas no. Instrumentales: si necesita apoyo en finanzas y cocina; lo demas no. Actividades laborales: complementarios medicos si, lo demas no. Discriminacion: ha sufrido discriminacion en algunos contextos, acoso laboral si, violencia fisica no, vulneracion de derechos no.",
        "Vida diaria: no requiere apoyo. Instrumentales: no aplica. Actividades laborales: no requiere apoyo. Discriminacion: no ha sufrido discriminacion.",
      ],
      prompt_fragment:
        "Si el profesional dice solo X o lo demas no, puedes marcar el resto en No. Si no lo dice, deja los no mencionados en null.",
    },
  },
};
