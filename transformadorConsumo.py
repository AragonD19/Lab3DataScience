import pandas as pd

# Instala openpyxl si no está instalado
try:
    import openpyxl
except ImportError:
    import os
    os.system('pip install openpyxl')

# Ruta del archivo .xlsx de entrada
input_excel_file = 'CONSUMO-2024-05.xlsx'

# Ruta del archivo .csv de salida
output_csv_file = 'Consumo.csv'

# Leer el archivo .xlsx especificando que la fila 7 es el encabezado (index 6 porque empieza desde 0)
excel_data = pd.read_excel(input_excel_file, header=6)

# Reemplazar los valores vacíos con None
excel_data = excel_data.where(pd.notnull(excel_data), None)

# Guardar los datos en un archivo .csv
excel_data.to_csv(output_csv_file, index=False)

print(f'Archivo convertido y guardado en {output_csv_file}')
