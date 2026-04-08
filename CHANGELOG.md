# Changelog

## 2.0.14 - 2026-04-07

Esta version consolida el rediseño operativo de `Seguimientos` y agrega protecciones para que el diligenciamiento diario sea mas claro y seguro para los profesionales.

### Cambios principales

- `SeguimientosWindow` y `SeguimientoEditorWindow` ahora muestran el flujo por etapas visibles, con lenguaje operativo en vez de nombres tecnicos de hojas.
- La ficha inicial se reorganiza en bloques mas claros, recupera un scroll mas uniforme y muestra la linea de tiempo de seguimientos en solo lectura.
- `Fecha fin contrato` ahora acepta texto libre y agrega una accion rapida para registrar `No aplica`.
- `Copiar datos del seguimiento anterior` conserva la fecha y los textos largos del seguimiento actual, pero copia el resto de campos operativos, incluidos asistentes.
- Los campos que van a sobreescribir datos ya diligenciados se resaltan en amarillo y `Guardar etapa` pide confirmacion explicita antes de reemplazarlos.
- Se agregan pruebas de workflow y runtime para cubrir el stage model, la navegacion por etapas, las confirmaciones de sobreescritura y el nuevo comportamiento de copiado.
- Se agrega la guia operativa [docs/seguimientos_flujo_operativo.md](docs/seguimientos_flujo_operativo.md) para adopcion del nuevo flujo.

### Como afecta a los usuarios

- El profesional ve con mayor claridad en que etapa va, que puede editar y cual es el siguiente paso sugerido.
- La ficha inicial y los seguimientos reducen carga cognitiva y evitan errores al corregir informacion ya existente.
- Cuando se intenta reemplazar un dato anterior, la app ahora lo hace visible antes de guardar.

## 2.0.13 - 2026-04-07

Esta version estabiliza el actualizador de Windows despues de los cambios recientes del handoff al instalador.

### Cambios principales

- Se bloquea el cierre de la ventana mientras el instalador esta corriendo para evitar interrupciones del proceso.
- Se eliminan artefactos auxiliares del updater que habian quedado despues de la iteracion anterior.
- El flujo del actualizador vuelve a una ruta mas simple y predecible para reducir riesgos operativos en release.

### Como afecta a los usuarios

- La actualizacion de la app es menos propensa a quedarse a medio camino por cierres prematuros o handoffs innecesariamente complejos.

## 2.0.12 - 2026-04-07

Esta version corrige una regresion en `Seguimientos` que podia dejar la ventana a medio renderizar y endurece la carga inicial de cédulas cuando la base de datos no responde o la sesión perdió permisos.

### Cambios principales

- `SeguimientosWindow` recupera su handler `_open_lsc_window`, evitando que el header falle al construir el botón `Solicitar Intérprete LSC` y que la vista quede incompleta.
- La búsqueda por cédula ahora muestra un mensaje visible cuando `usuarios_reca` no carga, deshabilita `Buscar` mientras no exista una lista válida y agrega `Recargar lista` para reintentar sin cerrar la ventana.
- Los errores `401/403`, de permisos o de conectividad al cargar cédulas se traducen a mensajes más accionables dentro del mismo flujo.
- Se agregan pruebas runtime para cubrir la regresión de render de `Seguimientos` y el reintento exitoso de recarga de cédulas.

### Como afecta a los usuarios

- `Seguimientos` vuelve a abrir completo en vez de quedar vacío o a medio construir por el fallo del botón de LSC.
- Si la lista de cédulas no carga por sesión, permisos o red, el usuario ya no queda con un combo vacío sin contexto y puede intentar la recarga desde la misma pantalla.

## 2.0.11 - 2026-04-07

Esta version corrige un fallo del actualizador que podia descargar el instalador correcto pero relanzar la app sin aplicar realmente la actualizacion.

### Cambios principales

- El watcher post-cierre del `updater` deja de ejecutar el instalador con `start /wait` y ahora lo invoca de forma directa desde el script temporal, evitando carreras con el relanzamiento.
- El handoff agrega un log del instalador en `%TEMP%` para facilitar diagnostico si Inno Setup vuelve a fallar en equipos de usuario final.
- Se agregan pruebas unitarias para validar el contenido del script post-cierre y asegurar que el instalador se ejecute sin `start /wait`.

### Como afecta a los usuarios

- Cuando aceptan actualizar, la instalacion ya no deberia volver a abrir la app antigua antes de terminar el setup.
- La version instalada deberia quedar realmente actualizada despues del relanzamiento.

## 2.0.10 - 2026-04-07

Esta version agrega un aviso automatico de actualizacion en el HUB para que los usuarios finales no pasen por alto que existe una version nueva.

### Cambios principales

- Se unifica en `app.py` la resolucion del estado de actualizacion en un snapshot reutilizable para el chequeo silencioso y la actualizacion manual.
- Al entrar al HUB, la app sigue consultando la version remota en segundo plano, actualiza el indicador `Version local | GitHub` y ahora muestra un prompt una sola vez por apertura cuando existe una version superior con instalador valido.
- Si el usuario acepta, se reutiliza exactamente el flujo existente de descarga, handoff al instalador y reinicio; si el usuario rechaza, la app continua normal sin volver a insistir en esa ejecucion.
- El boton manual `Actualizar aplicacion` ahora reutiliza el mismo snapshot y reporta un error claro si el release remoto no incluye el instalador esperado.
- Se agregan pruebas unitarias para el prompt automatico, el guard de una sola vez por ejecucion y el caso de release invalido sin asset de instalador.

### Como afecta a los usuarios

- Los usuarios reciben un aviso claro apenas entran al HUB cuando hay una version nueva disponible.
- Ya no dependen de notar por su cuenta el indicador de version o de abrir manualmente el flujo de actualizacion.
- Si prefieren seguir trabajando sin actualizar en ese momento, pueden hacerlo sin bloqueos ni prompts repetidos en la misma apertura de la app.

## 2.0.9 - 2026-04-06

Esta version incorpora el ajuste del formato maestro de `Condiciones de Vacante` despues de eliminar filas en la plantilla.

### Cambios principales

- Se corrige el mapeo de `section_2` para `herramientas_equipos`.
- Se recorren las referencias de `section_3`, `section_4`, `section_5`, `section_6`, `section_7` y `section_8` para alinearlas con el nuevo layout del Google Sheet.
- Se actualizan `SECTION_7_TITLE_ROW`, `SECTION_8_TITLE_ROW` y los offsets dinamicos usados al exportar observaciones y asistentes.

### Como afecta a los usuarios

- `Condiciones de Vacante` vuelve a escribir en las filas correctas del formato maestro actual.
- Las observaciones finales y la lista de asistentes ya no quedan desplazadas despues del cambio de plantilla.

## 2.0.8 - 2026-04-06

Esta version corrige el mecanismo de actualizacion para evitar bloqueos de antivirus durante el handoff al instalador.

### Cambios principales

- El `updater` deja de invocar `powershell.exe` en tiempo de ejecucion para consultar releases, descargar el instalador y esperar el cierre de la app.
- La instalacion diferida ahora usa un helper `.cmd` temporal para esperar el cierre, ejecutar el instalador y relanzar la aplicacion.
- Cuando falla la descarga del instalador, la app reporta un error claro y remite al release en vez de intentar fallbacks que disparaban proteccion conductual.

### Como afecta a los usuarios

- El flujo de actualizacion manual reduce el riesgo de ser bloqueado por soluciones como Norton.
- La app sigue pudiendo descargar el instalador y relanzarse, pero con un mecanismo menos agresivo para antivirus corporativos.

## 2.0.5 - 2026-04-06

Esta version corrige la reanudacion de formularios desde borradores y mejora la captura de detalles largos en `Evaluacion de Accesibilidad`.

### Cambios principales

- La seccion `1` vuelve a hidratar correctamente los datos de empresa desde cache o borrador y mantiene habilitado `Continuar` cuando el usuario regresa a corregir informacion antes de finalizar.
- `Seleccion Incluyente` y los demas formularios que reutilizan la seccion `1` comparten ahora una restauracion consistente del estado de empresa al reabrir o devolverse dentro del flujo.
- Los campos `Detalle` de `Evaluacion de Accesibilidad` ahora usan cajas de texto de `2` lineas con autoajuste hasta `10`, evitando cortes visuales cuando la respuesta es larga.
- Se agregaron pruebas de interfaz para validar la restauracion de seccion `1` y el comportamiento multilinea de los campos `Detalle`.

### Como afecta a los usuarios

- Un acta abierta desde `Borradores` ya no queda bloqueada en la seccion `1` despues de refrescar la empresa.
- Los usuarios pueden corregir informacion de empresa y continuar de nuevo hasta la seccion final sin perder la navegacion.
- En `Evaluacion de Accesibilidad`, los textos largos en `Detalle` se leen y editan mejor sin quedar comprimidos en una sola linea.

## 2.0.4 - 2026-04-06

Esta version ajusta la logica de dependencias en `Seleccion Incluyente` y corrige el mapeo de comunicacion escrita en la plantilla maestra.

### Cambios principales

- `4.2A` ahora deja independientes los terceros dropdowns de desplazamiento y ubicacion, sincronizando solo los dos primeros campos de esas preguntas.
- `4.2B` aplica reglas especificas desde el primer dropdown: `0` replica `0` al segundo y pone `No` en los auxiliares; `No aplica` replica `No aplica` a todos; `1`, `2` y `3` solo sincronizan el segundo y dejan en blanco los auxiliares.
- `comunicacion_escrita_nota` corrige su mapeo en Google Sheets para escribir en `O52`, conservando offsets correctos cuando hay multiples oferentes.
- Se agregaron pruebas para la nueva logica de sincronizacion y para validar el mapeo con offsets por bloque de oferente.

### Como afecta a los usuarios

- La seccion `4.2` ya no sobreescribe campos que deben diligenciarse de forma independiente.
- La nota de comunicacion escrita vuelve a caer en la celda correcta del formato maestro, tanto en procesos individuales como grupales.

## 2.0.3 - 2026-04-06

Esta version amplía la generación de PDFs para más actas, corrige un desajuste en inducción organizacional y endurece la reutilización de plantillas de Google Sheets en escenarios dinámicos.

### Cambios principales

- `Contratación Incluyente`, `Inducción Organizacional` e `Inducción Operativa` ahora retornan metadata completa de acta para habilitar la exportación a PDF con rotulación y metadatos consistentes.
- El flujo de finalización ahora encola exportación PDF para esos formularios adicionales, usando los mismos controles ya existentes para otros procesos.
- La publicación desde plantilla en Google Sheets ahora tolera mejor rangos que todavía no existen en pestañas nuevas, evitando falsos positivos al evaluar si una hoja ya estaba poblada.
- `Inducción Organizacional` corrige el número base de filas de asistentes para alinearse con la plantilla actual y evitar desplazamientos incorrectos.
- Se actualizó la documentación interna del proyecto y de módulos críticos para facilitar mantenimiento, debugging y releases futuros.

### Como afecta a los usuarios

- Más formularios generan su PDF final automáticamente después de publicar el acta, sin depender de pasos manuales adicionales.
- Los formularios con bloques dinámicos reutilizan mejor las plantillas remotas y reducen fallos por lectura de rangos fuera del grid actual.
- `Inducción Organizacional` queda alineada con la plantilla vigente y reduce riesgo de exportes con filas corridas en la sección de asistentes.

## 2.0.2 - 2026-04-03

Esta version corrige la distribucion de credenciales de Google en el instalador y evita que la interfaz confunda un problema de configuracion con falta de internet.

### Cambios principales

- El instalador ahora copia `service-account.json` al perfil del usuario desde la maquina de build, dejando lista la autenticacion de Google Drive y Google Sheets en instalaciones nuevas o actualizadas.
- El badge de estado del Hub ya no muestra `Sin conexión` cuando la red esta bien pero faltan credenciales o configuracion de servicios.
- Los estados degradados de servicios conectados ahora se clasifican como `Configuración incompleta`, `Credenciales inválidas` o `Servicios no disponibles` segun el tipo de falla.

### Como afecta a los usuarios

- Los formularios que dependen de Google dejan de fallar en equipos actualizados por ausencia del `service-account.json`.
- Cuando exista un problema de configuracion, el usuario vera un estado mas preciso y no un falso `Sin conexión`.

## 2.0.1 - 2026-04-03

Esta version corrige un faltante de empaquetado que podia impedir la finalizacion de formularios en instalaciones de usuario final.

### Cambios principales

- Se incluye `config.json` dentro del ejecutable distribuido y del instalador para que la app instalada pueda resolver IDs de Google Sheets y Google Drive definidos por configuracion.
- El proceso de build ahora valida antes y despues de PyInstaller que los recursos runtime criticos queden presentes en `dist\RECA_INCLUSION_LABORAL`.
- Se documento la ruta de recuperacion manual para instalaciones ya desplegadas que necesiten restaurar `config.json`.

### Como afecta a los usuarios

- `Presentacion Programa` y otros flujos que dependen de plantillas o carpetas remotas dejan de fallar por configuracion ausente en equipos instalados.
- Los nuevos releases tienen una proteccion adicional para evitar publicar instaladores incompletos.

## 2.0.0 - 2026-04-03

Esta version consolida una actualizacion amplia de estabilidad, seguridad operativa y experiencia de uso en los flujos principales de Inclusion Laboral.

### Cambios principales

- Se reforzo el inicio de sesion y la resolucion de perfil con autologin, fallbacks mas seguros y validaciones adicionales para reducir bloqueos al entrar a la app.
- Se endurecio el manejo de configuracion, secretos, cache y actualizaciones para evitar dependencias fragiles en entornos instalados y hacer mas segura la ruta de upgrade.
- Se reorganizo la retroalimentacion visual en formularios para mostrar errores y mensajes operativos mas claros antes de guardar o finalizar.
- `Seguimientos`, `Google Sheets` y `Drive` recibieron ajustes de resiliencia, compatibilidad y protecciones para reducir fallos intermitentes durante publicacion y sincronizacion.
- Se agregaron pruebas de seguridad, contrato de mensajes, runtime y UX para cubrir los cambios nuevos y prevenir regresiones en release.

### Como afecta a los usuarios

- El ingreso a la aplicacion y la continuidad de sesion son mas estables en equipos ya instalados.
- La publicacion de formatos y el proceso de actualizacion tienen mas validaciones antes de ejecutar cambios sobre recursos remotos.
- Los mensajes visibles durante errores o validaciones bloqueantes son mas claros y ayudan a corregir el problema sin perder contexto.

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
