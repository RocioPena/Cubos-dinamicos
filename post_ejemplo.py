from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import pandas as pd
import win32com.client
import pythoncom

app = FastAPI()

# ----- Configuración de conexión al cubo -----
CUBO = "SIS_2024"
CONNECTION_STRING = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    f"Initial Catalog={CUBO};"
    "Connect Timeout=60;"
)

# ----- Modelo de datos del request -----
class ConsultaDinamica(BaseModel):
    variables: list[str]
    unidades: list[str]
    fechas: list[str]
    filtros_where: list[str] | None = None

# ----- Función de consulta OLAP -----
def query_olap(connection_string: str, query: str) -> pd.DataFrame:
    pythoncom.CoInitialize()
    conn = win32com.client.Dispatch("ADODB.Connection")
    rs = win32com.client.Dispatch("ADODB.Recordset")
    conn.Open(connection_string)
    rs.Open(query, conn)
    
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

# ----- Endpoint de consulta avanzada -----
@app.post("/consulta_avanzada")
def consulta_dinamica(payload: ConsultaDinamica):
    try:
        # Construcción del MDX
        mdx = f"""
        SELECT 
            {{ {', '.join(payload.variables)} }} ON COLUMNS,
            {{ {', '.join(payload.unidades)} }} ON ROWS
        FROM [{CUBO}]
        WHERE ( {', '.join(payload.fechas)} {',' if payload.filtros_where else ''} {', '.join(payload.filtros_where or [])} )
        """

        df = query_olap(CONNECTION_STRING, mdx)
        return JSONResponse(content=df.to_dict(orient="records"))
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
