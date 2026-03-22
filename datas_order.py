import pandas as pd

# 1. Cargar la base de datos principal
# CORRECCIÓN: Se usa read_excel, no read_xlsx
try:
    df = pd.read_excel('Datos-Sabado-Noche.xlsx')
except Exception as e:
    print(f"Error al leer el archivo. Asegúrate de tener instalada la librería 'openpyxl' (pip install openpyxl). Error: {e}")
    exit()

# Nombres de columnas clave
COL_SEDE = 'university'
COL_ALIMENTACION = 'dietary_preference'
COL_DOC_TIPO = 'document_type'

# 2. Preparar el motor para escribir un Excel con varias hojas
ruta_salida = 'Reportes_Logistica_ENEUN.xlsx'
writer = pd.ExcelWriter(ruta_salida, engine='xlsxwriter')

# --- REPORTE 1: CONFIRMACIONES ---
# CORRECCIÓN: Agregamos .copy() para evitar la advertencia de Pandas
df_confirmados = df[df['confirm_submitted_at'].notna()].copy()

conf_sede = df_confirmados[COL_SEDE].value_counts().reset_index()
conf_sede.columns = ['Sede', 'Total Confirmados']

# Agregar fila de Total Nacional
total_nacional = pd.DataFrame([{'Sede': 'TOTAL NACIONAL', 'Total Confirmados': conf_sede['Total Confirmados'].sum()}])
conf_final = pd.concat([conf_sede, total_nacional], ignore_index=True)
conf_final.to_excel(writer, sheet_name='1_Confirmaciones', index=False)

# --- REPORTE 2: TIPO DE ALIMENTACION ---
alim_cruzado = pd.crosstab(df_confirmados[COL_SEDE], df_confirmados[COL_ALIMENTACION], margins=True, margins_name='TOTAL NACIONAL')
alim_cruzado.reset_index(inplace=True)
alim_cruzado.to_excel(writer, sheet_name='2_Alimentacion', index=False)

# --- REPORTE 3: MAYORIA DE EDAD (CENSO) ---
df_confirmados['Es_Mayor'] = df_confirmados[COL_DOC_TIPO].apply(lambda x: 'Mayor de Edad (CC)' if str(x).upper() == 'CC' else 'Menor de Edad / Otro')
edad_cruzado = pd.crosstab(df_confirmados[COL_SEDE], df_confirmados['Es_Mayor'], margins=True, margins_name='TOTAL NACIONAL')
edad_cruzado.reset_index(inplace=True)
edad_cruzado.to_excel(writer, sheet_name='3_Mayoria_Edad', index=False)

# --- REPORTE 4: REPRESENTANTES ---
# Manejo seguro en caso de que is_representative tenga valores nulos o texto en vez de booleanos
df_repres = df_confirmados[df_confirmados['is_representative'].isin([True, 'Sí', 'Si', 'TRUE', 1])]

repres_sede = df_repres[COL_SEDE].value_counts().reset_index()
repres_sede.columns = ['Sede', 'Total Representantes']

if not repres_sede.empty:
    total_rep = pd.DataFrame([{'Sede': 'TOTAL NACIONAL', 'Total Representantes': repres_sede['Total Representantes'].sum()}])
    repres_final = pd.concat([repres_sede, total_rep], ignore_index=True)
    repres_final.to_excel(writer, sheet_name='4_Representantes', index=False)
    
    # Lista detallada
    cols_repres = [c for c in ['first_name', 'last_name', COL_SEDE, 'representative_bodies', 'cuerpo_colegiado'] if c in df_repres.columns]
    df_repres[cols_repres].to_excel(writer, sheet_name='4_Detalle_Repres', index=False)

# --- REPORTE 5: CONDICIONES MEDICAS ---
df_medico = df_confirmados[df_confirmados['has_disability'].isin([True, 'Sí', 'Si']) | df_confirmados['allergies'].isin([True, 'Sí', 'Si'])]

cols_medico = [c for c in ['first_name', 'last_name', COL_SEDE, 'has_disability', 'disability_specify', 'allergies', 'allergy_specify'] if c in df_medico.columns]
if not df_medico.empty:
    df_medico_detalle = df_medico[cols_medico]
    df_medico_detalle.to_excel(writer, sheet_name='5_Condiciones_Medicas', index=False)

# Guardar y cerrar el archivo Excel
writer.close()

print(f"¡Proceso terminado exitosamente! Revisa el archivo: {ruta_salida}")