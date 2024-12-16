from datetime import datetime
import pandas as pd
import os

carpeta_destino = "media/sharepoint"
nuevo_nombre = "Horario_General_Copia_v2.xlsx"
carpeta_destino = os.path.abspath(carpeta_destino)

ruta_guardado = os.path.join(carpeta_destino, nuevo_nombre)
# Cargar la hoja específica para procesar
#hoja = '09-12 al 15-12'
hojas_seleccionadas = ['25-11 al 01-12', '02-12 al 08-12', '09-12 al 15-12']

 # Crear una lista para almacenar los datos extraídos
datos_extraidos = []

for hoja in hojas_seleccionadas:
    df = pd.read_excel(ruta_guardado, sheet_name=hoja, header=None)

    # Extraer los encabezados de los días (fechas) de la fila 0
    encabezados_dias = df.iloc[0, 2:].dropna().tolist()

    # Iterar por las filas para capturar nombres y turnos
    for i, row in df.iterrows():
        if i >= 11 and pd.notnull(row[1]):  # A partir de la fila 11
            nombre = row[1]

            # Iterar sobre las columnas correspondientes a los días
            for idx, encabezado in enumerate(encabezados_dias):
                turno_col = 2 + idx * 3  # Columna del turno
                turno = row[turno_col]

                # Guardar solo si hay un turno válido
                if pd.notnull(turno):
                    datos_extraidos.append({
                        'Fecha': encabezado,
                        'Nombre': nombre,
                        'Turno': turno
                    })

# Convertir los datos extraídos a un DataFrame
df_resultados = pd.DataFrame(datos_extraidos)
df_resultados['Solo Fecha'] = pd.to_datetime(df_resultados['Fecha'].str.extract(r'(\d{2}/\d{2}/\d{4})')[0],format='%d/%m/%Y').dt.strftime('%Y-%m-%d')

# Mostrar el resultado
df_resultados.head(10)

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

reporte_sharepoint = os.path.join(carpeta_destino, f'sharepoint_df_{timestamp}.xlsx')
df_resultados.to_excel(reporte_sharepoint, index=False, engine='openpyxl')

