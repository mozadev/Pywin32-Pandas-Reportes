import win32com.client
import os

def guardar_excel_como(carpeta_destino, nuevo_nombre):
    # Conectar con la aplicación de Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Mostrar Excel para ver el proceso (puedes poner False para ocultarlo)

    # Obtener el workbook activo
    workbook = excel.ActiveWorkbook

    carpeta_destino = os.path.abspath(carpeta_destino)

    if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino)

    # Ruta completa para guardar el archivo
    ruta_guardado = os.path.join(carpeta_destino, nuevo_nombre)

    try:
        # Guardar una copia en la carpeta especificada
        workbook.SaveAs(ruta_guardado)
        print(f"Archivo guardado en: {ruta_guardado}")

        workbook.Close(SaveChanges=False)

    except Exception as e:
        print(f"Error al guardar el archivo: {e}")
    finally:
        # Cerrar el workbook sin cerrar Excel
        # workbook.Close(SaveChanges=False)  # Descomenta si deseas cerrar el archivo después de guardar
        pass

# Ejemplo de uso
carpeta_destino = "media/sharepoint"
nuevo_nombre = "Horario_General_Copia_v2.xlsx"
guardar_excel_como(carpeta_destino, nuevo_nombre)
