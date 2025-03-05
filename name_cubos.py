import pandas as pd
import adodbapi

def rows_to_df(rows) -> pd.DataFrame:
    """ Convierte los resultados de una consulta de adodbapi a un DataFrame de Pandas. """
    df = pd.DataFrame(data=dict(zip(rows.columnNames.keys(), rows.ado_results)))
    return df

# Conexión al servidor OLAP (sin especificar Initial Catalog para evitar errores de acceso)
dsn_base = 'Provider=MSOLAP.8;Password=Temp123!;Persist Security Info=True;User ID=SALUD\DGIS15;'
dsn_base += 'Data Source=pwidgis03.salud.gob.mx;Update Isolation Level=2;'

try:
    # Conectar sin especificar una base de datos para obtener los catálogos disponibles
    conn = adodbapi.connect(dsn_base, timeout=600)
    cursor = conn.cursor()

    # Obtener los catálogos disponibles
    cursor.execute("SELECT [catalog_name] FROM $system.DBSCHEMA_CATALOGS")
    rows = cursor.fetchall()
    catalogos_df = rows_to_df(rows)
    print("Catálogos disponibles:")
    print(catalogos_df.to_string())  # Mostrar todos los catálogos
    conn.close()

    # Obtener los cubos para cada catálogo
    cubos_totales = []
    for catalogo in catalogos_df.iloc[:, 0]:
        print(f"Consultando cubos en catálogo: {catalogo}")
        
        # Establecer la conexión con el catálogo actual
        dsn = dsn_base + f'Initial Catalog={catalogo};'
        conn = adodbapi.connect(dsn, timeout=600)
        cursor = conn.cursor()
        
        cursor.execute("SELECT CUBE_NAME FROM $system.MDSCHEMA_CUBES")
        rows = cursor.fetchall()
        df_cubos = rows_to_df(rows)
        df_cubos['DATABASE'] = catalogo  # Agregar la base de datos de origen
        cubos_totales.append(df_cubos)
        print(df_cubos.to_string())  # Mostrar todos los cubos del catálogo
        
        conn.close()
    
    # Guardar los cubos en un CSV si hay demasiados para la terminal
    if cubos_totales:
        cubos_df = pd.concat(cubos_totales, ignore_index=True)
        cubos_df.to_csv('cubos_olap.csv', index=False, encoding='utf-8')
        print("Todos los cubos han sido guardados en cubos_olap.csv")
    
except Exception as e:
    print(f"Error de conexión: {e}")
