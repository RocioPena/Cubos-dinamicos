import pandas as pd
import adodbapi

def rows_to_df(rows) -> pd.DataFrame:
    """ 
    Convierte los resultados de una consulta de adodbapi a un DataFrame de Pandas.
    rows: resultado de la consulta
    """
    df = pd.DataFrame(data=dict(zip(rows.columnNames.keys(), rows.ado_results)))\
        .assign(_id=lambda x: range(len(x)))
    return df

# Conexión al cubo específico "Recursos"
cubo = 'Recursos'
dsn = 'Provider=MSOLAP.8;Password=Temp123!;Persist Security Info=True;User ID=SALUD\DGIS15;'
dsn += f'Data Source=pwidgis03.salud.gob.mx;Update Isolation Level=2;Initial Catalog={cubo};'

try:
    conn = adodbapi.connect(dsn, timeout=600)
    cursor = conn.cursor()
    
    # Obtener los encabezados del cubo "Recursos"
    print(f"\nConsultando los encabezados del cubo '{cubo}'")
    cursor.execute("SELECT * FROM $SYSTEM.MDSCHEMA_MEASURES")
    rows = cursor.fetchall()
    df_encabezados = rows_to_df(rows)
    
    print("Encabezados del cubo 'Recursos':")
    print(df_encabezados.to_string())
    
    conn.close()
    
except Exception as e:
    print(f"Error de conexión: {e}")
