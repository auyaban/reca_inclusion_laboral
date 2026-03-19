# Section 2 Labs Voice Contract

Especificacion humana para `seleccion_incluyente_labs -> section_2`.

Objetivo:
- Un audio por oferente y por subseccion.
- El modelo debe devolver JSON estructurado con claves exactas.
- El modelo no debe inventar datos ni completar negativos si el profesional no los dijo de forma explicita.

Reglas generales:
- `numero` y `edad` quedan en `null` en el modelo. El desktop los resuelve localmente.
- Si el profesional no da un dato con claridad, usar `null`.
- Si el profesional dice `solo`, `unicamente`, `ninguna`, `ninguno`, `lo demas no`, `todas no`, `todo no`, se permite propagacion explicita de `No`.
- Si el profesional enumera positivos sin cerrar con una negacion explicita del resto, los no mencionados quedan `null`.
- Los textos narrativos van en `desarrollo_actividad` o en los campos `*_nota`.

## 2. Datos del oferente

Campos:
- `nombre_oferente`: nombre completo del oferente.
- `cedula`: numero de cedula en texto.
- `certificado_porcentaje`: porcentaje del certificado, por ejemplo `45` o `45.5`.
- `discapacidad`: valor exacto del dropdown.
- `telefono_oferente`: telefono principal.
- `resultado_certificado`: valor exacto del dropdown.
- `cargo_oferente`: cargo o vacante.
- `nombre_contacto_emergencia`
- `parentesco`
- `telefono_emergencia`
- `fecha_nacimiento`: preferir `dd/mm/aaaa` si el audio lo permite.
- `pendiente_otros_oferentes`: valor exacto del dropdown.
- `lugar_firma_contrato`
- `fecha_firma_contrato`
- `cuenta_pension`: valor exacto del dropdown.
- `tipo_pension`: valor exacto del dropdown.

Dropdowns exactos:
- `discapacidad`
  - `Discapacidad visual pérdida total de la visión`
  - `Discapacidad visual baja visión`
  - `Discapacidad auditiva`
  - `Discapacidad auditiva hipoacusia`
  - `Trastorno de espectro autista`
  - `Discapacidad intelectual`
  - `Discapacidad física`
  - `Discapacidad física usuario en silla de ruedas`
  - `Discapacidad psicosocial`
  - `Discapacidad múltiple`
  - `No aplica`
- `resultado_certificado`
  - `Aprobado`
  - `No aprobado`
  - `Pendiente`
- `pendiente_otros_oferentes`
  - `Si`
  - `No`
  - `Por Confirmar`
- `cuenta_pension`
  - `Si`
  - `No`
  - `Por Confirmar`
- `tipo_pension`
  - `Pension Invalidez`
  - `Subsidiada`
  - `Especial de vejez`
  - `Victimas conflicto`
  - `Familiar`
  - `Regimen especial`
  - `No aplica`

Variantes esperadas:
- `por confirmar`, `aun no confirma`, `pendiente de confirmar` -> `Por Confirmar`
- `silla de ruedas`, `usuario de silla de ruedas` -> `Discapacidad física usuario en silla de ruedas`
- `autismo`, `TEA` -> `Trastorno de espectro autista`
- `baja vision` -> `Discapacidad visual baja visión`

Contraejemplos:
- No escribir `aprobado con observaciones`. Debe ser `Aprobado`, `No aprobado` o `Pendiente`.
- No escribir `pension invalidez`. Debe ser `Pension Invalidez`.

## 3. Desarrollo de la actividad

Campo:
- `desarrollo_actividad`: relato corrido y claro. Puede tener varias frases.

Buenas respuestas esperadas:
- Resumen de la entrevista o reunion.
- Observaciones relevantes.
- Acuerdos o pendientes.

Contraejemplos:
- No partirlo en listas artificiales.
- No llenar dropdowns desde este bloque si el audio solo es narrativo.

## 4.1 Condiciones medicas y de salud

Para `section_4_1_salud`, la salida debe incluir ademas una capa semantica en `semantic.section_4_1_health`.

Estados semanticos esperados:
- `medications.support_level`: `none` | `low` | `medium` | `high` | `not applicable`
- `medications.status`: `not taking` | `self managed` | `third party managed` | `unknown` | `not applicable`
- `medications.schedule_status`: `self managed` | `third party managed` | `unknown` | `not applicable`
- `allergies.support_level`: `none` | `low` | `medium` | `high` | `not applicable`
- `allergies.status`: `none reported` | `self managed` | `unknown` | `described` | `not applicable`
- `restrictions.support_level`: `none` | `low` | `medium` | `high` | `not applicable`
- `restrictions.status`: `none reported` | `self managed` | `unknown` | `does not know management` | `not applicable`
- `specialist_controls.support_level`: `none` | `low` | `medium` | `high` | `not applicable`
- `specialist_controls.attendance`: `attends and self manages` | `attends` | `unknown` | `not applicable`
- `specialist_controls.frequency`: `monthly` | `quarterly` | `semiannual` | `other` | `not applicable`
- En los cuatro bloques, `details` guarda el detalle corto clinico.

Regla operativa:
- `semantic.section_4_1_health` es la fuente principal de interpretacion.
- `candidate` puede ir en `null` cuando el dropdown exacto no sea totalmente obvio; el desktop hace la traduccion final.

### Medicamentos
- `medicamentos_nivel_apoyo`
  - `0. No requiere apoyo.`
  - `1. Nivel de apoyo Bajo.`
  - `2. Nivel de apoyo medio.`
  - `3. Nivel de apoyo alto.`
  - `No aplica.`
- `medicamentos_conocimiento`
  - `1. Conoce los medicamentos que consume.`
  - `2. Un tercero es quien conoce los medicamentos que consume.`
  - `3. No conoce los medicamentos que consume.`
  - `No aplica.`
  - `0. No requiere apoyo.`
- `medicamentos_horarios`
  - `1. Conoce los horarios de toma de medicamentos que consume.`
  - `2. Es un tercero quien conoce los horarios de la toma de medicamentos.`
  - `3. No conoce los horarios de toma de medicamentos que consume.`
  - `0. No requiere apoyo.`
  - `No aplica.`
- `medicamentos_nota`: nombre del medicamento o aclaracion breve.

### Alergias
- `alergias_nivel_apoyo`
  - `0. No requiere apoyo.`
  - `1. Nivel de apoyo Bajo.`
  - `2. Nivel de apoyo medio.`
  - `3. Nivel de apoyo alto.`
  - `No aplica.`
- `alergias_tipo`
  - `0. No presenta alergias.`
  - `1. Presenta alergias y sabe darle manejo.`
  - `2. No conoce si presenta alguna alergia.`
  - `3. Presenta alergias a: medicamentos, sustancias y productos quimicos, alimentos, animales, entre otros.`
  - `No aplica.`
- `alergias_nota`

### Restricciones medicas
- `restriccion_nivel_apoyo`
  - `0. No requiere apoyo.`
  - `1. Nivel de apoyo Bajo.`
  - `2. Nivel de apoyo medio.`
  - `3. Nivel de apoyo alto.`
  - `No aplica.`
- `restriccion_conocimiento`
  - `0. No tiene restricciones medicas.`
  - `1. Tiene restricciones medicas y conoce su manejo.`
  - `2. No conoce si tiene restricciones medicas.`
  - `3. Si tiene restricciones medicas y desconoce su manejo.`
  - `No aplica.`
- `restriccion_nota`

### Controles con especialista
- `controles_nivel_apoyo`
  - `0. No requiere apoyo.`
  - `1. Nivel de apoyo Bajo.`
  - `2. Nivel de apoyo medio.`
  - `3. Nivel de apoyo alto.`
  - `No aplica.`
- `controles_asistencia`
  - `No aplica.`
  - `2. Si asiste a controles medicos con especialista.`
  - `3. No sabe si tiene controles medicos con especialista.`
  - `1. Asiste a controles medicos con especialista y conoce el manejo.`
  - `0. No requiere apoyo.`
- `controles_frecuencia`
  - `Mensual`
  - `Trimestral`
  - `Semestral`
  - `Otra frecuencia`
  - `No aplica`
- `controles_nota`

Variantes esperadas:
- `no sabe` -> opcion de desconocimiento.
- `cada mes` -> `Mensual`
- `cada tres meses` -> `Trimestral`
- `cada seis meses` -> `Semestral`

## 4.2A Habilidades basicas de la vida diaria

### Desplazamiento
- `desplazamiento_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `desplazamiento_modo`
  - `0. Se desplaza de manera independiente sin necesidad de apoyos (ortesis, baston, silla de ruedas entre otros).`
  - `1. Se desplaza de forma independiente con un apoyo temporal (ortesis, baston, silla de ruedas entre otros).`
  - `2. Se desplaza de manera independiente con un apoyo permanente (ortesis, baston, silla de ruedas entre otros).`
  - `3. No se desplaza de manera independiente. Requiere el acompanamiento de un tercero y un apoyo (ortesis, baston, silla de ruedas entre otros).`
  - `No aplica.`
- `desplazamiento_transporte`
  - `Caminando.`
  - `Bicicleta.`
  - `Transmilenio, Sitp.`
  - `Vehiculo propio.`
  - `Vehiculo especial.`
  - `No aplica.`
- `desplazamiento_nota`

### Ubicacion
- `ubicacion_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `ubicacion_ciudad`
  - `0. Sabe ubicarse en la ciudad de manera autonoma.`
  - `1. Sabe ubicarse en la ciudad pero haciendo uso de aplicaciones (Maps, Waze, entre otros).`
  - `2. Requiere de acompanamiento para ubicarse.`
  - `3. No sabe ubicarse en la ciudad.`
- `ubicacion_aplicaciones`
  - `Se ubica por puntos de referencia y direcciones.`
  - `No se ubica por puntos de referencia.`
  - `Se ubica por puntos cardinales.`
  - `No aplica`
- `ubicacion_nota`

### Dinero
- `dinero_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `dinero_reconocimiento`
  - `Autonomo.`
  - `Con apoyo familiar.`
- `dinero_manejo`
  - `0. Reconoce y maneja el dinero de manera autonoma.`
  - `1. Reconoce y maneja el dinero pero en ocasiones requiere apoyo.`
  - `2. Solo reconoce el dinero pero no lo sabe manejar.`
  - `3. No reconoce el dinero y no lo sabe manejar.`
  - `No aplica.`
- `dinero_medios`
  - `Dinero fisico, plastico y digital.`
  - `Dinero fisico y plastico.`
  - `Dinero fisico.`
  - `Dinero plastico y digital.`
  - `Dinero plastico.`
  - `Dinero digital.`
  - `Dinero digital y fisico.`
- `dinero_nota`

### Presentacion personal
- `presentacion_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `presentacion_personal`
  - `0. Su codigo de vestuario es acorde al contexto.`
  - `1. Su codigo de vestuario es acorde al contexto, pero presenta oportunidades de mejora.`
  - `2. Su codigo de vestuario es medianamente acorde al contexto.`
  - `3. Su codigo de vestuario no es acorde al contexto.`
  - `No aplica.`
- `presentacion_nota`

### Comunicacion escrita
- `comunicacion_escrita_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `comunicacion_escrita_apoyo`
  - `0. Si conoce y maneja los apoyos (Jaws, Magic, el lector de pantalla de Windows/IOS).`
  - `1. Maneja algunos apoyos de comunicacion escrita, pero no todos en general.`
  - `2. Conoce pero no maneja apoyos.`
  - `3. No conoce, ni maneja los apoyos.`
  - `No aplica.`
- `comunicacion_escrita_nota`

### Comunicacion verbal
- `comunicacion_verbal_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `comunicacion_verbal_apoyo`
  - `0. Si conoce y maneja los apoyos (Centro de relevo, entre otros).`
  - `1. Maneja algunos apoyos, pero no los conoce todos en general (Centro de relevo, entre otros).`
  - `2. Conoce pero no maneja apoyos.`
  - `3. No conoce, ni maneja los apoyos.`
  - `No aplica.`
- `comunicacion_verbal_nota`

### Toma de decisiones
- `decisiones_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `toma_decisiones`
  - `0. Toma las decisiones de manera autonoma.`
  - `1. Toma decisiones pero en ocasiones requiere el apoyo de un tercero.`
  - `2. Debe consultar con un tercero para la toma de decisiones.`
  - `3. Requiere el apoyo de un tercero para tomar decisiones.`
  - `No aplica.`
- `toma_decisiones_nota`

## 4.2B Actividades, apoyos y discriminacion

### Regla central para subitems Si/No/No aplica
- Si el profesional dice `solo alimentacion; lo demas no`, se marca `aseo_alimentacion = Si` y el resto del subgrupo en `No`.
- Si el profesional dice `si necesita apoyo en finanzas y cocina` pero no dice `lo demas no`, solo se marcan esas en `Si` y las otras quedan `null`.
- Si el profesional dice `no aplica` para todo el bloque o subgrupo, usa `No aplica` donde exista.

### Vida diaria
- `aseo_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `alimentacion`
  - `0. No requiere apoyo en sus actividades de la vida diaria.`
  - `1. Requiere apoyo en algunas actividades de la vida diaria.`
  - `2. Requiere apoyo en la mayoria de actividades de la vida diaria.`
  - `3. Requiere apoyo en todas las actividades de la vida diaria.`
  - `No aplica.`
- `aseo_criar_apoyo`, `aseo_comunicacion_apoyo`, `aseo_ayudas_apoyo`, `aseo_alimentacion`, `aseo_movilidad_funcional`, `aseo_higiene_aseo`
  - `Si`
  - `No`
  - `No aplica`
- `aseo_nota`

### Instrumentales
- `instrumentales_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `instrumentales_actividades`
  - `0. No requiere apoyo en actividades instrumentales de la vida diaria.`
  - `1. Requiere apoyo en algunas actividades instrumentales de la vida diaria.`
  - `2. Requiere apoyo en la mayoria de actividades instrumentales de la vida diaria.`
  - `3. Requiere apoyo en todas las actividades instrumentales de la vida diaria.`
  - `No aplica.`
- `instrumentales_criar_apoyo`, `instrumentales_comunicacion_apoyo`, `instrumentales_movilidad_apoyo`, `instrumentales_finanzas`, `instrumentales_cocina_limpieza`, `instrumentales_crear_hogar`, `instrumentales_salud_cuenta_apoyo`
  - `Si`
  - `No`
  - `No aplica`
- `instrumentales_nota`

### Actividades laborales
- `actividades_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `actividades_apoyo`
  - `0. No requiere apoyo en sus actividades laborales.`
  - `1. Requiere apoyo en algunas actividades laborales.`
  - `2. Requiere apoyo en la mayoria de actividades laborales.`
  - `3. Requiere apoyo en todas las actividades laborales.`
  - `No aplica`
- `actividades_esparcimiento_apoyo`, `actividades_esparcimiento_cuenta_apoyo`, `actividades_complementarios_apoyo`, `actividades_complementarios_cuenta_apoyo`, `actividades_subsidios_cuenta_apoyo`
  - `Si`
  - `No`
  - `No aplica`
- `actividades_nota`

### Discriminacion
- `discriminacion_nivel_apoyo`: mismos niveles `0-3` o `No aplica.`
- `discriminacion`
  - `0. No ha sufrido de discriminacion.`
  - `1. Ha sufrido de discriminacion en algunos contextos.`
  - `2. Ha sufrido de discriminacion en repetidas ocasiones.`
  - `3. Ha sufrido de discriminacion a los largo del ciclo vital.`
  - `No aplica.`
- `discriminacion_violencia_apoyo`, `discriminacion_violencia_cuenta_apoyo`, `discriminacion_vulneracion_apoyo`, `discriminacion_vulneracion_cuenta_apoyo`
  - `Si`
  - `No`
  - `No aplica`
- `discriminacion_nota`

Variantes esperadas:
- `lo demas no`, `las demas no`, `solo estas`, `unicamente estas`
- `no requiere apoyo`
- `no aplica`
- `no sabe`

Contraejemplos:
- No marcar automaticamente `No` en items no mencionados si el profesional no cerro explicitamente el resto.
- No asumir que `no recuerda` es igual a `No`. En ese caso debe ir `null` o la opcion de desconocimiento si existe.
