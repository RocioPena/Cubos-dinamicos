# Consumo de Cubos Dinámicos

Este proyecto tiene como objetivo conectarse y consumir datos desde cubos OLAP de SQL Server Analysis Services (SSAS) a través de una API desarrollada en Python.

### Entorno de trabajo para ejecutar la API

- Sistema operativo: Windows  
- Python: 3.12  
- DB: SQL Server  
- Driver: MSOLAP

Es necesario instalar las [Bibliotecas cliente de Analysis Services](https://learn-microsoft-com.translate.goog/en-us/analysis-services/client-libraries?view=asallproducts-allversions&_x_tr_sl=en&_x_tr_tl=es&_x_tr_hl=es&_x_tr_pto=tc), según la versión que corresponda a tu entorno. A continuación se muestra una imagen referencial:

![Bibliotecas cliente de Analysis Services](image.png)

Una vez instaladas las bibliotecas, se deben instalar las siguientes dependencias ejecutando los siguientes comandos en la terminal:

```bash
pip install requirements.txt      # Instalar librerias.
```


Para ejecutar la API:

```bash
python -m uvicorn api:app --host 0.0.0.0 --port 8080
python -m uvicorn api_cubos:app --host 0.0.0.0 --port 8080
```
### Entorno de trabajo para consumir la API
 - Sistema operativo: Multiplataforma
 - Python: 3.12

Para consumir la API 
```bash
python consumo_api.py
```

Solo que se tiene que poner que enpoint se quiere consumir y la IP de la API que se esta ejecutando como servidor