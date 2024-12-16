import pandas as pd
import xlrd

# download_path = os.path.abspath("media/semaforo/")

#semaforo_path = "media/semaforo/reporte_semaforo_2024-11-01_2024-11-30_20241212205218.xls"
semaforo_path = "media/semaforo/reporte_semaforo_2024-11-01_2024-11-30_20241212212728.xls"
newcallcenter_path = "media/newcallcenter/new__call_center2024-11-01_2024-11-30_20241212205258.xlsx"
#semaforo_df = pd.read_excel(semaforo_path, engine="xlrd" )

workbook = xlrd.open_workbook(semaforo_path, ignore_workbook_corruption=True)
semaforo_df = pd.read_excel(workbook)

newcallcenter_df = pd.read_excel(newcallcenter_path, skiprows=6)

print(semaforo_df.head())
print(newcallcenter_df.head())