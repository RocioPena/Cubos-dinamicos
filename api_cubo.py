from fastapi import FastAPI
from fastapi.responses import JSONResponse
import pandas as pd
import win32com.client
import pythoncom
import os

app = FastAPI()

# Configuración de conexión OLAP
CUBO = "SIS_2024"
CONNECTION_STRING = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "Update Isolation Level=2;"
    f"Initial Catalog={CUBO};"
)

# Ejecutar consulta y devolver DataFrame
def query_sql(connection_string: str, query: str) -> pd.DataFrame:
    pythoncom.CoInitialize()
    conn = win32com.client.Dispatch("ADODB.Connection")
    rs = win32com.client.Dispatch("ADODB.Recordset")

    conn.Open(connection_string)
    rs.Open(query, conn)

    columns = [rs.Fields.Item(i).Name for i in range(rs.Fields.Count)]
    data = []

    while not rs.EOF:
        row = [rs.Fields.Item(i).Value for i in range(rs.Fields.Count)]
        data.append(row)
        rs.MoveNext()

    rs.Close()
    conn.Close()
    pythoncom.CoUninitialize()

    return pd.DataFrame(data, columns=columns)


@app.get("/dim_variables")
def get_dim_variables():
    """Consulta la dimensión DIM VARIABLES y devuelve los resultados"""
    try:
        query = "SELECT * FROM [SIS_2024].[$DIM VARIABLES]"
        df = query_sql(CONNECTION_STRING, query)
        return df.to_dict(orient="records")
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.get("/exportar_dim_variables")
def export_dim_variables():
    """Consulta la dimensión DIM VARIABLES y guarda archivo Excel"""
    try:
        query = "SELECT * FROM [SIS_2024].[$DIM VARIABLES]"
        df = query_sql(CONNECTION_STRING, query)
        os.makedirs("consulta_cubos", exist_ok=True)
        path = "consulta_cubos/dim_variables.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")
        return {"status": "ok", "archivo": path}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
