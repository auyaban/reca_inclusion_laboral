# Changelog

## 1.2.7 - 2026-03-30

Esta version endurece el flujo de `Seguimientos` frente a fallas temporales de red y deja mas seguro el proceso de actualizacion.

### Cambios principales

- Se corrigio el manejo de errores transitorios al leer el estado de casos de `Seguimientos`, evitando popups crudos por fallas como `WinError 10053`.
- La apertura del editor de `Seguimientos` ahora precarga la estructura del caso en segundo plano y muestra errores amigables si Drive o Google Sheets fallan temporalmente.
- Se redujo la agresividad del probe periodico de Drive para evitar ruido innecesario de conectividad en segundo plano.
- Se elimino codigo duplicado en `drive_upload.py` y se fortalecio la clasificacion entre errores transitorios y errores permanentes.
- Se agregaron pruebas de resiliencia para carga de casos, apertura del editor y manejo de errores de transporte.

### Como afecta a los usuarios

- Buscar una cedula y abrir un caso en `Seguimientos` ahora es mas estable cuando la red esta intermitente.
- Las fallas temporales de Drive dejan de verse como errores tecnicos de Windows y pasan a mostrarse como mensajes operativos.
- La app solo ofrecera la actualizacion cuando exista una version realmente superior, evitando instalaciones ambiguas sobre el mismo tag.

## 1.2.5 - 2026-03-27

Esta version corrige un riesgo serio de perdida de datos en formularios de induccion y agrega una forma segura de reabrir formularios ya terminados.

### Cambios principales

- Se corrigio el autoguardado de `Induccion Organizacional` y `Induccion Operativa` para que no vacie `section_3` al navegar entre secciones.
- El autoguardado ahora protege datos ya guardados frente a payloads vacios o sospechosos y mantiene historial local por seccion para recuperacion.
- La finalizacion y el cierre ahora se bloquean si una seccion obligatoria queda vacia despues de haber tenido datos, evitando exportes incompletos.
- Se agrego un indicador visible de `Ultimo guardado` dentro del flujo de formularios.
- El boton `Labs` del hub fue reemplazado por `Terminados`.
- `Terminados` permite reabrir en el flujo normal los formularios finalizados de los ultimos 30 dias con los datos precargados.
- Se agrego `tzdata` a dependencias para empaquetar correctamente la zona horaria en la app instalada.

### Como afecta a los usuarios

- Se reduce de forma importante el riesgo de perder informacion diligenciada al cambiar de seccion o finalizar un formulario.
- Si un formulario ya finalizado debe reabrirse, el usuario puede hacerlo desde `Terminados` sin volver a digitar la informacion.
- La app instalada mantiene mejor consistencia de fechas y horas entre entornos Windows.

## 1.2.2 - 2026-03-25

Esta version corrige errores de cierre y acceso que estaban afectando a usuarios en la app instalada.

### Cambios principales

- Los formularios ahora resuelven templates y recursos desde el bundle instalado, no solo desde rutas de desarrollo.
- `Seleccion Incluyente` y `Contratacion Incluyente` ahora toleran variaciones razonables en nombres de templates grupales para evitar fallos por nombres no exactos.
- La generacion de archivos Excel ahora endurece nombres y rutas de salida para evitar errores de Windows como `WinError 3`.
- El inicio de sesion ya no se bloquea si la consulta de perfil sobre `profesionales` falla por permisos; la app entra con un perfil minimo degradado.

### Como afecta a los usuarios

- Los formatos que antes no finalizaban por templates no encontrados vuelven a poder cerrarse desde la app instalada.
- Los cierres que fallaban por rutas invalidas o demasiado largas en Windows ahora tienen fallback automatico.
- Los usuarios autenticados ya no quedan bloqueados al iniciar sesion por una restriccion de lectura sobre `profesionales`.

## 1.2.1 - 2026-03-19

Esta version hace un ajuste pequeno de interfaz en `Seleccion Incluyente`.

### Cambios principales

- En la seccion de oferentes, el campo visible `Resultado certificado` ahora se muestra como `Resultado`.

### Como afecta a los usuarios

- No cambia el funcionamiento del formato, el mapeo a Excel ni la informacion guardada.
- Solo simplifica el texto visible en pantalla para que coincida mejor con el uso diario.

## 1.2.0 - 2026-03-19

Esta version consolida una actualizacion grande de los flujos de Inclusion Laboral. El foco estuvo en hacer mas estable el trabajo diario, mejorar el soporte para procesos grupales y dejar mas confiable el diligenciamiento por voz donde ya estaba habilitado.

### Cambios principales

- `Seleccion Incluyente` ahora soporta correctamente formatos individuales y grupales desde el mismo flujo productivo.
- `Contratacion Incluyente` ahora soporta formatos individuales y grupales para 2 o mas vinculados, incluyendo offsets para grupos mas grandes.
- Se corrigio el mapeo de Excel seccion por seccion para que los datos caigan en la celda correcta en los templates individuales y grupales.
- `Desarrollo de la actividad` ahora se comporta de forma dinamica en UI cuando el proceso pasa de una persona a varias, sin perder la informacion escrita.
- Se fortalecio la exportacion a Excel y la generacion del instalador para que el proceso de release siga funcionando con los nuevos templates.
- Se ajusto la sincronizacion con `usuarios_reca` para evitar errores por shapes inconsistentes y por cedulas duplicadas dentro del mismo envio.
- El buscador de empresa en seccion 1 ahora filtra en tiempo real y permite cargar la informacion al seleccionar la empresa, sin depender del boton de busqueda.
- Se fortalecio el flujo experimental de voz en `Seleccion Incluyente Labs` y `Condiciones de Vacante Labs`, con mejor interpretacion semantica y mejor autollenado.

### Como afecta a los usuarios

- Si trabajas con un solo oferente o vinculado, el formato sigue viendose como antes, pero con menos errores de ubicacion en Excel.
- Si trabajas con varios oferentes o vinculados, ya no necesitas usar plantillas manuales aparte: el sistema ajusta el formato grupal automaticamente.
- El cierre del formato y la generacion del archivo final quedan mas confiables, especialmente cuando hay varios registros en una misma sesion.
- En los flujos con dictado, el sistema entiende mejor respuestas en lenguaje natural y reduce la cantidad de correcciones manuales.

### Nota operativa

- Los formatos grupales nuevos dependen de los templates actualizados incluidos en esta version.
- Las colas antiguas que hayan quedado guardadas localmente con payloads viejos pueden seguir reintentando hasta limpiarse, pero los envios nuevos ya salen con la logica corregida.
