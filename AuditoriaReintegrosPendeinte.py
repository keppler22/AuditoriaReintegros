import os
import xlsxwriter
import pandas as pd

# Configuración de Pandas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

sheet_dic = {}

# Verifica la existencia del archivo en la ruta especifica
file_path = '/content/PIAM_UNICAUCA.xlsx'
file_path1 = '/content/ReintegrosPendendientes.xlsx'
file_path2 = '/content/PIAM_REINTEGROS_UNICAUCA.xlsx'
resultado_path = '/content/ReintegrosPendeintesAuditado.xlsx'

for path in [file_path, file_path1]:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"{path} no encontrado.")
    else:
        print(f"Archivo {path} encontrado.")

# Abre los archivos en modo binario para verificar problemas de acceso
for path in [file_path, file_path1, file_path2]:
    try:
        with open(path, 'rb') as f:
            print(f"Archivo {path} abierto satisfactoriamente en modo binario.")
    except OSError as e:
        print(f"Error al abrir el archivo {path}: {e}")


reintegrosPendientes = pd.read_excel(file_path1)
sq100824 = pd.read_excel(file_path, sheet_name='SQ010824') 
reintegros231 = pd.read_excel(file_path2, sheet_name='RELACIONREINTEGROS23-1') 

df_reintegros = pd.merge(reintegrosPendientes, sq100824[['Id  factura', 'Documento','Periodico Academico','Nombre de Destino']], left_on='ID FACTURA', right_on='Id  factura', how='left')
df_reintegrosfinal = pd.merge(df_reintegros, reintegros231[['ID FACTURA', 'ESTADO','TELEFONO','CELULAR','FONDO ']], left_on='ID FACTURA', right_on='ID FACTURA', how='left')

columnas_reporte = ['ID',	'CODG',	'NOMBRE COMPLETO',	'Nombre de Destino',	'VALOR REINTEGRO',	'Periodico Academico',	'ESTADO', 'FONDO ',	'TELEFONO',	'CELULAR_y','EMAIL_INSTITUCIONAL']
df_reintegrosfinalordenado = df_reintegrosfinal[columnas_reporte]
df_reintegrosfinalordenado.to_excel(resultado_path, index=False)
print(f"Archivo de resultado guardado en {resultado_path}")
