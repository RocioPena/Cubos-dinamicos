from fastapi import FastAPI, Body
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import pandas as pd
import win32com.client
import pythoncom

app = FastAPI()

# Configuración de conexión
CUBO = "SIS_2024"
CONNECTION_STRING = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "Update Isolation Level=2;"
    f"Initial Catalog={CUBO};"
    "Connect Timeout=60;"
)

# Modelo para recibir filtros
class ConsultaRequest(BaseModel):
    variables: list[str]
    unidades: list[str]
    fechas: list[str]

# Funciones para ejecutar MDX

def ejecutar_mdx(connection_string: str, mdx_query: str) -> pd.DataFrame:
    pythoncom.CoInitialize()
    conn = win32com.client.Dispatch("ADODB.Connection")
    rs = win32com.client.Dispatch("ADODB.Recordset")
    conn.Open(connection_string)
    rs.Open(mdx_query, conn)

    fields = [rs.Fields.Item(i).Name for i in range(rs.Fields.Count)]
    data = []
    while not rs.EOF:
        row = [rs.Fields.Item(i).Value for i in range(rs.Fields.Count)]
        data.append(row)
        rs.MoveNext()

    rs.Close()
    conn.Close()
    pythoncom.CoUninitialize()

    return pd.DataFrame(data, columns=fields)

@app.post("/consulta_avanzada")
def consulta_dinamica(filtros: ConsultaRequest = Body(...)):
    try:
        # Armar elementos MDX
        mdx_variables = ", ".join([f"[Measures].[{v}]" for v in filtros.variables])
        mdx_unidades = ", ".join([f"[DIM MODULO].[Módulo].[{u}]" for u in filtros.unidades])
        mdx_fechas = ", ".join([f"[DIM TIEMPO].[Año].[{f}]" for f in filtros.fechas])

        # MDX Final
        mdx_query = f"""
        SELECT
            {{ {mdx_variables} }} ON COLUMNS,
            CROSSJOIN({{ {mdx_unidades} }}, {{ {mdx_fechas} }}) ON ROWS
        FROM [{CUBO}]
        """

        df = ejecutar_mdx(CONNECTION_STRING, mdx_query)
        return df.to_dict(orient="records")

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
