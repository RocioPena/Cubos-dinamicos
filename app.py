import pandas as pd
import adodbapi

def rows_to_df(rows) -> pd.DataFrame:
    """ 
    Convierte los resultados de una consulta de adobdapi a un DataFrame de Pandas.
    rows: resultado de la consulta
    """
    df = pd.DataFrame(data=dict(zip(rows.columnNames.keys(), rows.ado_results)))\
        .assign(_id=lambda x: range(len(x)))
    return df

cubo = 'SIS_2024'
conn = adodbapi.connect('Provider=MSOLAP.8;Password=Temp123!;Persist Security Info=True;User ID=SALUD\DGIS15;'
                        f'Data Source=pwidgis03.salud.gob.mx;Update Isolation Level=2;Initial Catalog={cubo}', timeout=600)
cursor = conn.cursor()
#cursor.execute("""SELECT [catalog_name] FROM $system.DBSCHEMA_CATALOGS""")
#rows = cursor.fetchall()
#catalogos = list(rows)

#for catalogo in catalogos:
#	print(catalogo)

#cursor.execute("""
#SELECT [CATALOG_NAME] as [DATABASE],
#CUBE_NAME AS [CUBE], DIMENSION_CAPTION AS [DIMENSION]
#FROM $system.MDSchema_Dimensions
#WHERE DIMENSION_CAPTION='Measures'
#""")
#rows = cursor.fetchall()
#columnas = list(rows)
#print(rows_to_df(rows))

#for columna in columnas:
#	print(columna)

#cursor.execute("""
#SELECT *  
#FROM $SYSTEM.MDSCHEMA_MEASURES  
#""")
#rows = cursor.fetchall()
#columnas = list(rows)
#print(rows_to_df(rows))

#cursor.execute("""
#SELECT *
#FROM [SIS_2024].[$DIM TIEMPO]
#""")
#rows = cursor.fetchall()
#columnas = list(rows)
#print(rows_to_df(rows))

#for columna in columnas:
#	print(columna)

cursor.execute("""
SELECT *
FROM $system.MDSCHEMA_MEMBERS
""")
rows = cursor.fetchall()
df = rows_to_df(rows)
df.to_csv('output.csv', sep=',', encoding='utf-8')

# cursor.execute("""
# SELECT *
# FROM [SIS_2024].[Measures]
# WHERE [totales].[$dim unidad.entidad]="HIDALGO" AND
# ([totales].[$dim variables.clave] = "BIO03" OR
# [totales].[$dim variables.clave] = "BIO50" OR
# [totales].[$dim variables.clave] = "VBC01" OR
# [totales].[$dim variables.clave] = "VBC02" OR
# [totales].[$dim variables.clave] = "VBC03" OR
# [totales].[$dim variables.clave] = "VBC51" OR
# [totales].[$dim variables.clave] = "VBC52" OR
# [totales].[$dim variables.clave] = "VBC53" OR
# [totales].[$dim variables.clave] = "VBC54" OR
# [totales].[$dim variables.clave] = "VBC55" OR
# [totales].[$dim variables.clave] = "VAC06" OR
# [totales].[$dim variables.clave] = "VHB01" OR
# [totales].[$dim variables.clave] = "VHB02" OR
# [totales].[$dim variables.clave] = "VHB03" OR
# [totales].[$dim variables.clave] = "VHB04" OR
# [totales].[$dim variables.clave] = "VHB05" OR
# [totales].[$dim variables.clave] = "VHB06" OR
# [totales].[$dim variables.clave] = "VHB51" OR
# [totales].[$dim variables.clave] = "VHB52" OR
# [totales].[$dim variables.clave] = "VHB53" OR
# [totales].[$dim variables.clave] = "VHB54" OR
# [totales].[$dim variables.clave] = "VHB55" OR
# [totales].[$dim variables.clave] = "BIO88" OR
# [totales].[$dim variables.clave] = "VAC87" OR
# [totales].[$dim variables.clave] = "VHA51" OR
# [totales].[$dim variables.clave] = "VHA52" OR
# [totales].[$dim variables.clave] = "VAC12" OR
# [totales].[$dim variables.clave] = "VAC13" OR
# [totales].[$dim variables.clave] = "VPD51" OR
# [totales].[$dim variables.clave] = "VPD52" OR
# [totales].[$dim variables.clave] = "VAC17" OR
# [totales].[$dim variables.clave] = "VAC18" OR
# [totales].[$dim variables.clave] = "VAC19" OR
# [totales].[$dim variables.clave] = "VAC93" OR
# [totales].[$dim variables.clave] = "VAC94" OR
# [totales].[$dim variables.clave] = "VNC01" OR
# [totales].[$dim variables.clave] = "VNC02" OR
# [totales].[$dim variables.clave] = "VNC03" OR
# [totales].[$dim variables.clave] = "VNC51" OR
# [totales].[$dim variables.clave] = "VNC52" OR
# [totales].[$dim variables.clave] = "VNC53" OR
# [totales].[$dim variables.clave] = "VNC54" OR
# [totales].[$dim variables.clave] = "VNC55" OR
# [totales].[$dim variables.clave] = "VNC56" OR
# [totales].[$dim variables.clave] = "VNC57" OR
# [totales].[$dim variables.clave] = "VNC58" OR
# [totales].[$dim variables.clave] = "VNP01" OR
# [totales].[$dim variables.clave] = "VNP51" OR
# [totales].[$dim variables.clave] = "VAC23" OR
# [totales].[$dim variables.clave] = "VAC81" OR
# [totales].[$dim variables.clave] = "VTV01" OR
# [totales].[$dim variables.clave] = "VTV02" OR
# [totales].[$dim variables.clave] = "VTV03" OR
# [totales].[$dim variables.clave] = "VTV51" OR
# [totales].[$dim variables.clave] = "VTV52" OR
# [totales].[$dim variables.clave] = "VTV53" OR
# [totales].[$dim variables.clave] = "VTV54" OR
# [totales].[$dim variables.clave] = "VTV55" OR
# [totales].[$dim variables.clave] = "VAC82" OR
# [totales].[$dim variables.clave] = "VAC83" OR
# [totales].[$dim variables.clave] = "VAC91" OR
# [totales].[$dim variables.clave] = "VDV51" OR
# [totales].[$dim variables.clave] = "VDV52" OR
# [totales].[$dim variables.clave] = "VDV53" OR
# [totales].[$dim variables.clave] = "VAC84" OR
# [totales].[$dim variables.clave] = "VAC85" OR
# [totales].[$dim variables.clave] = "VAC92" OR
# [totales].[$dim variables.clave] = "VPH05" OR
# [totales].[$dim variables.clave] = "VPH06" OR
# [totales].[$dim variables.clave] = "VPH07" OR
# [totales].[$dim variables.clave] = "VPH08" OR
# [totales].[$dim variables.clave] = "VPH09" OR
# [totales].[$dim variables.clave] = "VPH10" OR
# [totales].[$dim variables.clave] = "VPH11" OR
# [totales].[$dim variables.clave] = "VPH51" OR
# [totales].[$dim variables.clave] = "VPH52" OR
# [totales].[$dim variables.clave] = "VPH53" OR
# [totales].[$dim variables.clave] = "VPH54" OR
# [totales].[$dim variables.clave] = "VPH55" OR
# [totales].[$dim variables.clave] = "VPH56" OR
# [totales].[$dim variables.clave] = "VPH57" OR
# [totales].[$dim variables.clave] = "VPH58" OR
# [totales].[$dim variables.clave] = "VPH59" OR
# [totales].[$dim variables.clave] = "VPH60" OR
# [totales].[$dim variables.clave] = "VAC36" OR
# [totales].[$dim variables.clave] = "VAC38" OR
# [totales].[$dim variables.clave] = "VAR01" OR
# [totales].[$dim variables.clave] = "VAR51" OR
# [totales].[$dim variables.clave] = "VAR52" OR
# [totales].[$dim variables.clave] = "VAR53" OR
# [totales].[$dim variables.clave] = "VAC39" OR
# [totales].[$dim variables.clave] = "VAC40" OR
# [totales].[$dim variables.clave] = "VAC43" OR
# [totales].[$dim variables.clave] = "VAC46" OR
# [totales].[$dim variables.clave] = "VAC47" OR
# [totales].[$dim variables.clave] = "VAC48" OR
# [totales].[$dim variables.clave] = "VAC51" OR
# [totales].[$dim variables.clave] = "VAC54" OR
# [totales].[$dim variables.clave] = "VAC55" OR
# [totales].[$dim variables.clave] = "VAC56" OR
# [totales].[$dim variables.clave] = "VAC59" OR
# [totales].[$dim variables.clave] = "VAC62" OR
# [totales].[$dim variables.clave] = "VTD01" OR
# [totales].[$dim variables.clave] = "VTD02" OR
# [totales].[$dim variables.clave] = "VTD03" OR
# [totales].[$dim variables.clave] = "VTD05" OR
# [totales].[$dim variables.clave] = "VTD07" OR
# [totales].[$dim variables.clave] = "VTD09" OR
# [totales].[$dim variables.clave] = "VTD11" OR
# [totales].[$dim variables.clave] = "VTD13" OR
# [totales].[$dim variables.clave] = "VTD14" OR
# [totales].[$dim variables.clave] = "VTD16" OR
# [totales].[$dim variables.clave] = "VTD19" OR
# [totales].[$dim variables.clave] = "VTD20" OR
# [totales].[$dim variables.clave] = "VTD21" OR
# [totales].[$dim variables.clave] = "VTD22" OR
# [totales].[$dim variables.clave] = "VTD23" OR
# [totales].[$dim variables.clave] = "VTD24" OR
# [totales].[$dim variables.clave] = "VTD25" OR
# [totales].[$dim variables.clave] = "VTD26" OR
# [totales].[$dim variables.clave] = "VTD27" OR
# [totales].[$dim variables.clave] = "VTD28" OR
# [totales].[$dim variables.clave] = "VTD29" OR
# [totales].[$dim variables.clave] = "VTD30" OR
# [totales].[$dim variables.clave] = "VTD31" OR
# [totales].[$dim variables.clave] = "VTD32" OR
# [totales].[$dim variables.clave] = "VTD33" OR
# [totales].[$dim variables.clave] = "VTD34" OR
# [totales].[$dim variables.clave] = "VTD35" OR
# [totales].[$dim variables.clave] = "VTD36" OR
# [totales].[$dim variables.clave] = "VTD51" OR
# [totales].[$dim variables.clave] = "VTD52" OR
# [totales].[$dim variables.clave] = "VTD53" OR
# [totales].[$dim variables.clave] = "VTD54" OR
# [totales].[$dim variables.clave] = "VTD55" OR
# [totales].[$dim variables.clave] = "VTD56" OR
# [totales].[$dim variables.clave] = "VTD57" OR
# [totales].[$dim variables.clave] = "VTD58" OR
# [totales].[$dim variables.clave] = "VTD59" OR
# [totales].[$dim variables.clave] = "VTD60" OR
# [totales].[$dim variables.clave] = "VTD61" OR
# [totales].[$dim variables.clave] = "VTD62" OR
# [totales].[$dim variables.clave] = "VTD63" OR
# [totales].[$dim variables.clave] = "VTD64" OR
# [totales].[$dim variables.clave] = "VTD65" OR
# [totales].[$dim variables.clave] = "VTD66" OR
# [totales].[$dim variables.clave] = "VTD67" OR
# [totales].[$dim variables.clave] = "VTD68" OR
# [totales].[$dim variables.clave] = "VTD69" OR
# [totales].[$dim variables.clave] = "VTD70" OR
# [totales].[$dim variables.clave] = "VTD71" OR
# [totales].[$dim variables.clave] = "VTD72" OR
# [totales].[$dim variables.clave] = "VTD73" OR
# [totales].[$dim variables.clave] = "VTD74" OR
# [totales].[$dim variables.clave] = "VTD75" OR
# [totales].[$dim variables.clave] = "VTD76" OR
# [totales].[$dim variables.clave] = "VTD77" OR
# [totales].[$dim variables.clave] = "VTD78" OR
# [totales].[$dim variables.clave] = "VTD79" OR
# [totales].[$dim variables.clave] = "VTD80" OR
# [totales].[$dim variables.clave] = "VTD81" OR
# [totales].[$dim variables.clave] = "VTD82" OR
# [totales].[$dim variables.clave] = "VAC63" OR
# [totales].[$dim variables.clave] = "VDP01" OR
# [totales].[$dim variables.clave] = "VDP51" OR
# [totales].[$dim variables.clave] = "VRV01" OR
# [totales].[$dim variables.clave] = "VRV02" OR
# [totales].[$dim variables.clave] = "VRV03" OR
# [totales].[$dim variables.clave] = "VRV04" OR
# [totales].[$dim variables.clave] = "VRV51" OR
# [totales].[$dim variables.clave] = "VRV52" OR
# [totales].[$dim variables.clave] = "VRV53" OR
# [totales].[$dim variables.clave] = "VRV54" OR
# [totales].[$dim variables.clave] = "VAC67" OR
# [totales].[$dim variables.clave] = "VAC68" OR
# [totales].[$dim variables.clave] = "VAC69" OR
# [totales].[$dim variables.clave] = "VAC70" OR
# [totales].[$dim variables.clave] = "VHX01" OR
# [totales].[$dim variables.clave] = "VHX02" OR
# [totales].[$dim variables.clave] = "VHX03" OR
# [totales].[$dim variables.clave] = "VHX04" OR
# [totales].[$dim variables.clave] = "VHX51" OR
# [totales].[$dim variables.clave] = "VHX52" OR
# [totales].[$dim variables.clave] = "VHX53" OR
# [totales].[$dim variables.clave] = "VHX54" OR
# [totales].[$dim variables.clave] = "VHX55" OR
# [totales].[$dim variables.clave] = "VHX56" OR
# [totales].[$dim variables.clave] = "VHX57" OR
# [totales].[$dim variables.clave] = "VHX58"
# ) 
# """)
# rows = cursor.fetchall()
# resultados = list(rows)

# for resultado in resultados:
# 	print(resultado)