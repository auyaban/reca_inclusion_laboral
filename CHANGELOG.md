# Changelog

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
