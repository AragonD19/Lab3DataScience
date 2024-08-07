import pandas as pd

# Instala openpyxl si no está instalado
try:
    import openpyxl
except ImportError:
    import os
    os.system('pip install openpyxl')

# Ruta del archivo .xlsx de entrada
input_excel_file = 'Precios-Promedio-Nacionales-Diarios-2024-3.xlsx'

# Definir las filas de encabezado y las filas a eliminar para cada hoja
header_and_remove_rows = {
    '2024': {'header': 7, 'remove': [8]},
    '2023': {'header': 7, 'remove': [8]},
    '2022': {'header': 6, 'remove': [7]},
    '2021': {'header': 6, 'remove': [7]}
}

# Leer el archivo .xlsx y obtener todas las hojas
excel_data = pd.read_excel(input_excel_file, sheet_name=None)

# Crear un DataFrame vacío para almacenar todas las hojas combinadas
combined_data = pd.DataFrame()

# Iterar sobre cada hoja y concatenarla al DataFrame combinado
for sheet_name, data in excel_data.items():
    if sheet_name in header_and_remove_rows:
        config = header_and_remove_rows[sheet_name]
        header_row = config['header']
        rows_to_remove = config['remove']
        
        # Leer solo las filas necesarias para los encabezados
        sheet_data = pd.read_excel(input_excel_file, sheet_name=sheet_name, header=None)
        
        # Seleccionar la fila de encabezado
        headers = sheet_data.iloc[header_row]
        sheet_data = sheet_data.iloc[header_row + 1:]
        sheet_data.columns = headers
        
        # Eliminar filas específicas
        sheet_data = sheet_data[~sheet_data.index.isin(rows_to_remove)]
        
        # Eliminar las últimas 3 filas
        sheet_data = sheet_data.iloc[:-3]
        
    else:
        # Leer la hoja con la primera fila como encabezado por defecto
        sheet_data = pd.read_excel(input_excel_file, sheet_name=sheet_name)
    
    # Reemplazar los valores vacíos con None
    sheet_data = sheet_data.where(pd.notnull(sheet_data), None)
    
    # Opcional: agregar una columna con el nombre de la hoja para referencia
    sheet_data['Sheet'] = sheet_name
    
    # Concatenar los datos al DataFrame combinado
    combined_data = pd.concat([combined_data, sheet_data], ignore_index=True)

# Ruta del archivo .csv de salida
output_csv_file = 'Precios.csv'

# Guardar los datos combinados en un archivo .csv
combined_data.to_csv(output_csv_file, index=False)

print(f'Todas las hojas combinadas y guardadas en {output_csv_file}')
