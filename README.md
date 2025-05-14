# Generador de Contratos

Esta aplicación permite generar múltiples contratos a partir de una plantilla de Word utilizando datos de un archivo Excel. Tiene un modelo de contrato y un excel.

## Uso

1. Archivo excel: debe contener las siguientes columnas
   - nombre
   - apellido
   - cedula
   - num_casas
   - prestamo

2. Preparar plantilla de word con los siguientes tags
   - {{nombre}}
   - {{apellido}}
   - {{cedula}}
   - {{num_casas}}
   - {{prestamo}}

3. En la interfaz gráfica:
   - Seleccionar el archivo excel con los datos
   - Seleccionar la plantilla de Word
   - Elegir la carpeta donde se guardarán los contratos generados
   - Hacer clic en "Generar Contratos"
