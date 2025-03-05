import pandas as pd

def clean_pivot_table(file_path, sheet_name='Hoja1', output_path='prueda.xlsx'):
    # Cargar el archivo Excel
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Identificar la fila donde realmente comienzan los encabezados
    header_row = df[df.iloc[:, 0] == 'Entidad'].index[0]
    
    # Extraer encabezados y limpiar valores NaN propagando hacia adelante
    headers = df.iloc[header_row].fillna(method='ffill')
    
    # Crear un DataFrame limpio sin las filas superiores innecesarias
    df_clean = df.iloc[header_row+1:].copy()
    df_clean.columns = headers  # Asignar encabezados
    
    # Rellenar valores fusionados en las columnas clave
    df_clean.fillna(method='ffill', inplace=True)
    
    # Guardar el resultado en un nuevo archivo Excel
    df_clean.to_excel(output_path, index=False)
    print(f'Datos limpiados y guardados en: {output_path}')

# Uso
document_path = "Cubos dinamicos 2024.xlsx"
clean_pivot_table(document_path)
