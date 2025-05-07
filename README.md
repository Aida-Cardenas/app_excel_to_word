# Generador de Contratos

Esta aplicación permite generar múltiples contratos a partir de una plantilla de Word utilizando datos de un archivo Excel.

## Requisitos

- Python 3.8 o superior
- Las dependencias listadas en `requirements.txt`

## Instalación

1. Clona este repositorio
2. Instala las dependencias:
```bash
pip install -r requirements.txt
```

## Uso

1. Prepara tu archivo Excel con las siguientes columnas:
   - nombre
   - apellido
   - cedula
   - num_casas
   - prestamo

2. Prepara tu plantilla de Word con los siguientes marcadores:
   - {{nombre}}
   - {{apellido}}
   - {{cedula}}
   - {{num_casas}}
   - {{prestamo}}

3. Ejecuta la aplicación:
```bash
python app.py
```

4. En la interfaz gráfica:
   - Selecciona el archivo Excel con los datos
   - Selecciona la plantilla de Word
   - Elige la carpeta donde se guardarán los contratos generados
   - Haz clic en "Generar Contratos"

Los contratos generados se guardarán en la carpeta seleccionada con el nombre `contrato_[cedula].docx`. 