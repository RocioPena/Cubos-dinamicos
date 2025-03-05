# Cubos-dinamicos a XLSX?

El archivo convertir.py identifica las columnas fusionadas y luego generar un script en Python 3.11 que las "desfusione" correctamente.

1. Carga el archivo Excel y detecta la fila donde comienzan los encabezados.
2. Desfusiona columnas propagando valores hacia adelante donde sea necesario.
3. Elimina filas irrelevantes y asigna nombres de columnas correctos.
4. Guarda el archivo limpio en un nuevo Excel llamado cleaned_data.xlsx.

### Requisitos de instalacion
```
pip install pandas

pip install openpyxl
```




#                                      #
#            🚀 Sección de cubos      #
#                                      #


En ```name_cubos.py``` muestra el catalogo de cubos

|id| catalog_name|
|--|--|
|0    |                  CLUES_2019|
|1      |              CLUES_2019_C|
|2       |       CONAPO_NACIMIENTOS|
|3       |     Cubo solo sinba 2020|
|4     |       Cubo solo sinba 2021|
|5  |          Cubo solo sinba 2022|
6|            Cubo solo sinba 2023|
7 |             cubo_lesiones_2017|
8  |            cubo_lesiones_2018|
9   |            Cubo_Lesiones2019|
10   |          Cubo_pobla_2019_DH|
11    |           CuboLesiones2020|
12     |          CuboLesiones2021|
13      |         CuboLesiones2022|
14       |        CuboLesiones2023|
15        |       CuboLesiones2024|
16         |      CuboLesiones2025|
17          |        CuboSec_18_20|
18           |       CuboSec_18_21|
19            |      CuboSec_18_22|
|20|            CuboSectorial_18_19|
21  |          CuboSectorial_18_20|
22   |            Defunciones_hist|
23    |        DEFUNCIONES_PC_2020|
24     |          DERECHOHABIENCIA|
25      |              Detecciones|
26   |                     EGRESOS|
27  | egresos_procedimientos_sinba|
28    |    egresos_productos_sinba|
29     |             egresos_sinba|
30      |              EGRESOS2018|
31       |          Egresos2019_19|
32        |            Egresos2020|
33         |           Egresos2021|
34          |          Egresos2022|
35           |         Egresos2023|
36            |        Egresos2024|
37             |       Egresos2025|
38              |     IND_DEM_PROY|
39               |        Lesiones|
40|             LESIONES_Historico|
41 |                lesiones_SINBA|
42  |                Maternas_2019|
43   |               Maternas_2020|
44    |           NACIMIENTOS_2018|
45     |          NACIMIENTOS_2019|
46      |         NACIMIENTOS_2020|
47       |        NACIMIENTOS_2021|
48        |       NACIMIENTOS_2022|
49         |      NACIMIENTOS_2023|
50          |     NACIMIENTOS_2024|
51           |    NACIMIENTOS_2025|
52            |     Personal_Salud|
53|            Pob_2018_aseg_nov15|
54 |          POB_MIT_PROYECCIONES|
55  |       POB_MIT_PROYECCIONES_C|
56   |                   Poblacion|
57    |             PROCEDIMIENTOS|
58     |       PROCEDIMIENTOS_2019|
59      |            PROY_DEF_EDAD|
60       |         PROY_DEF_EDAD_C|
61        |               Recursos|
62         |  recursos_sector_ok_2|
63          |       Reporte_Diario|
64           |    saeh_sector_hist|
65            |           saeh2011|
66             |          saeh2012|
67              |         saeh2013|
68               |        saeh2016|
69                |       saeh2017|
70                 |      saeh2018|
71                  |     saeh2019|
72                   |SALUD_MENTAL|
73|                SALUD_MENTAL_18|
74 |                      seed2013|
75  |                     seed2016|
76   |             seed2017_cierre|
77    |                  SICUENTAS|
78     |                sinac_2015|
79      |               sinac_2016|
80       |        SINAC_SINBA_2017|
81        |              sinac2017|
82         |             SINERHIAS|
83          |      SINERHIAS2025_1|
84           |        SIS_2017_NEW|
85            |      SIS_2018_NEW2|
86             |        SIS_2019_2|
87              |         SIS_2024|
88               |        SIS_2025|
89                |  SIS_SECTORIAL|
90|                SIS_SECTORIAL_C|
91 |                       sis2010|
92  |                      sis2011|
93   |                     sis2012|
94    |                    SIS2014|
95     |                   sis2015|
96      |                  sis2016|
97       |    TEF_NAC_PROYECCIONES|
98        | TEF_NAC_PROYECCIONES_C|
99         |       Urgencias_SINBA|
100         |        urgencias2011|
101          |       urgencias2012|
102           |      urgencias2013|
103            |     URGENCIAS2017|
104             |    URGENCIAS2018|
105              |   Urgencias2019|
106               |  Urgencias2020|
107                | Urgencias2021|
108                 |Urgencias2022|
109|                 Urgencias2023|
110 |                Urgencias2024|
111  |               Urgencias2025|



Obtener los encabezados de los cubos con ```contenido_cubos.py```

| catalog_name | schema_name | cube_name           | measure_name          | measure_unique_name            | measure_caption     | measure_guid | measure_aggregator | data_type | numeric_precision | numeric_scale | measure_units | description | expression | measure_is_visible | levels_list                                      | measure_name_sql_column_name | measure_unqualified_caption | measuregroup_name    | measure_display_folder | default_format_string | _id |
|-------------|------------|---------------------|-----------------------|--------------------------------|---------------------|--------------|-------------------|-----------|------------------|--------------|---------------|-------------|------------|-------------------|------------------------------------------------|-----------------------------|----------------------------|---------------------|------------------------|----------------------|-----|
| Recursos    | None       | $Cat Clues2017      | Cat Clues2017         | [Measures].[Cat Clues2017]     | Cat Clues2017      | None         | 1                 | 1         | 10               | -1           | None          | None        | True       | [Cat Clues2017].[CLUES].[Nombre]           | Cat Clues2017                        | None                          | Cat Clues2017MG      |                        |                      | 0   |
| Recursos    | None       | $Cat Mpo Refer      | Cat Mpo Refer        | [Measures].[Cat Mpo Refer]     | Cat Mpo Refer      | None         | 1                 | 1         | 10               | -1           | None          | None        | True       | None                                         | Cat Mpo Refer                        | None                          | Cat Mpo ReferMG      |                        |                      | 1   |
| Recursos    | None       | $Cat Tipo Uni2006   | Cat Tipo Uni2006     | [Measures].[Cat Tipo Uni2006]  | Cat Tipo Uni2006   | None         | 1                 | 1         | 10               | -1           | None          | None        | True       | [Cat Tipo Uni2006].[Tipo de unidad].[Tipo de unidad] | Cat Tipo Uni2006                        | None                          | Cat Tipo Uni2006MG   |                        |                      | 2   |
| Recursos    | None       | $Cat Variables      | Cat Variables        | [Measures].[Cat Variables]     | Cat Variables      | None         | 1                 | 1         | 10               | -1           | None          | None        | True       | [Cat Variables].[Variable].[Variable]       | Cat Variables                        | None                          | Cat VariablesMG      |                        |                      | 3   |
| Recursos    | None       | $Year               | Year                 | [Measures].[Year]              | Year               | None         | 1                 | 1         | 10               | -1           | None          | None        | True       | [Year].[Año].[Año]                         | Year                                  | None                          | YearMG               |                        |                      | 4   |
| Recursos    | None       | Recursos            | Número               | [Measures].[Número]            | Número             | None         | 1                 | 5         | 16               | -1           | None          | None        | True       | None                                         | Número                                | Número                        | Fact Recursos2017    |                        |                      | 5   |
| Recursos    | None       | Recursos            | Unidades reportadas  | [Measures].[Unidades reportadas] | Unidades reportadas | None         | 1                 | 5         | 16               | -1           | None          | None        | True       | None                                         | Unidades reportadas                 | Unidades reportadas          | Fact Recursos2017    |                        |                      | 6   |
