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