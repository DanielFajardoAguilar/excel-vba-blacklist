# Excel/VBA – Blacklist Email Cleaner (Demo)

Macro en Excel + VBA que depura una lista de correos contra una blacklist. La blacklist puede contener correos completos y/o dominios en formato @dominio. El resultado se genera en un libro nuevo con una hoja de depuración y otra de resumen.

Confidencialidad:
El desarrollo original se utilizó con información de una empresa (bases internas y correos reales). Por lo mismo, no se publican datos reales. Este repositorio contiene una versión demo con datos ficticios y lógica equivalente para fines de portafolio.

## Para qué sirve
Cuando se requiere limpiar una base de correos antes de un envío (CRM, campañas, comunicación masiva), esta macro permite:
- Identificar correos que aparecen en blacklist
- Bloquear por correo exacto o por dominio (@dominio)
- Conservar el archivo original sin modificaciones
- Generar un libro nuevo con la lista depurada y un resumen de resultados

## Cómo funciona
1. Seleccionas el rango de correos a depurar (una sola columna, sin encabezado).
2. Seleccionas el rango de blacklist (una sola columna, sin encabezado).
   - La blacklist admite:
     - Correo completo: cliente01@empresa-demo.com
     - Dominio: @spam-demo.com
3. La macro evalúa cada registro y crea un libro nuevo con:
   - Depuracion: correos válidos (no bloqueados)
   - Resumen_Depuracion: métricas, lista de eliminados y top repetidos

## Estructura del repositorio
- src/ contiene el módulo VBA exportado (.bas)
- data/ contiene archivos de entrada demo (datos ficticios)
- output/ contiene ejemplos del archivo resultante e incluye capturas del input/output

## Uso
1. Abrir los archivos de entrada en data/
2. Importar src/DepurarPorBlacklist.bas en el Editor VBA (Alt + F11)
3. Ejecutar la macro DepurarPorBlacklist
4. Seleccionar los rangos solicitados (correos y blacklist)

## Salida
Se genera un libro nuevo con dos hojas:
- Depuracion
- Resumen_Depuracion

El resumen incluye:
- Registros evaluados y eliminados
- Emails eliminados (unicos)
- Registros repetidos eliminados (por duplicados)
- Lista de correos eliminados
- Top repetidos (filas eliminadas por email)

## Supuestos y notas
- Se normalizan correos (trim y minusculas) para evitar problemas por mayusculas/minusculas.
- Si un dominio se incluye como @dominio, cualquier correo con ese dominio se bloquea.
- Para fines de demo, los datos son ficticios.
