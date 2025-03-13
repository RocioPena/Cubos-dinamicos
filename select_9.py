import adodbapi
import pandas as pd

# Configuración de la conexión (ajusta esto según tu servidor)
conn_str = "Provider=SQLOLEDB;Data Source=TU_SERVIDOR;Initial Catalog=SIS_2024;Integrated Security=SSPI;"
conn = adodbapi.connect(conn_str)
cursor = conn.cursor()

# Ejecutar la consulta para $DIM UNIDAD
cursor.execute("""
SELECT * 
FROM [SIS_2024].[$DIM UNIDAD]
""")

# Obtener los datos
rows = cursor.fetchall()

# Convertir a DataFrame
df = pd.DataFrame(rows, columns=[desc[0] for desc in cursor.description])

# Mostrar el DataFrame
import ace_tools as tools
tools.display_dataframe_to_user(name="DIM UNIDAD", dataframe=df)

# Cerrar conexión
cursor.close()
conn.close()
