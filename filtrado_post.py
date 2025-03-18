from fastapi import FastAPI
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import pandas as pd
import win32com.client
import pythoncom
import math

app = FastAPI()

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

class ConsultaDinamica(BaseModel):
    variables: list[str]
    unidades: list[str]
    fechas: list[str]
    filtros_where: list[str] | None = None
    entidad_filtro: str | None = None  # NUEVO

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

def sanitize_result(data):
    if isinstance(data, float) and (math.isnan(data) or data == float("inf") or data == float("-inf")):
        return None
    elif isinstance(data, list):
        return [sanitize_result(x) for x in data]
    elif isinstance(data, dict):
        return {k: sanitize_result(v) for k, v in data.items()}
    return data

@app.post("/consulta_avanzada")
def consulta_dinamica(payload: ConsultaDinamica):
    try:
        if payload.entidad_filtro:
            filtro_mdx = (
                f"FILTER(\n"
                f"  {{ [DIM UNIDAD].[Entidad Municipio Localidad].[Entidad Municipio Localidad].MEMBERS }},\n"
                f"  [DIM UNIDAD].[Entidad Municipio Localidad].CurrentMember.Properties(\"Entidad\") = \"{payload.entidad_filtro}\"\n"
                f")"
            )
            mdx = (
                "SELECT "
                f"{{ {', '.join(payload.variables)} }} ON COLUMNS,\n"
                f"{{ {filtro_mdx} }} ON ROWS\n"
                f"FROM [{CUBO}]\n"
                f"WHERE ( {', '.join(payload.fechas)}"
                + (", " + ", ".join(payload.filtros_where) if payload.filtros_where else "")
                + ")"
            )
        else:
            mdx = (
                "SELECT "
                f"{{ {', '.join(payload.variables)} }} ON COLUMNS,\n"
                f"{{ {', '.join(payload.unidades)} }} ON ROWS\n"
                f"FROM [{CUBO}]\n"
                f"WHERE ( {', '.join(payload.fechas)}"
                + (", " + ", ".join(payload.filtros_where) if payload.filtros_where else "")
                + ")"
            )

        df = query_olap(CONNECTION_STRING, mdx)
        sanitized_data = sanitize_result(df.to_dict(orient="records"))
        return JSONResponse(content=sanitized_data)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
