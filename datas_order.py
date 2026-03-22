import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Panel Logística ENEUN", layout="wide")

st.title("Panel de Control Logístico y Censos")

try:
    df = pd.read_excel('Datos-Sabado-Noche.xlsx')
except Exception as e:
    st.error(f"Error al cargar el archivo: {e}")
    st.stop()

df['Es_Mayor'] = df['document_type'].apply(lambda x: 'Mayor de Edad' if str(x).strip().upper() == 'CC' else 'Menor de Edad')

def extraer_campamento(respuesta):
    texto = str(respuesta).lower()
    if 'campamento' in texto and ('sí' in texto or 'si' in texto or 'true' in texto):
        return 'Acampa'
    return 'No Acampa'

def extraer_direccion(respuesta):
    texto = str(respuesta).lower()
    if 'direccion' in texto or 'dirección' in texto:
        return str(respuesta)
    return 'No especifica'

df['Campamento'] = df['answers'].apply(extraer_campamento)
df['Direccion_Hospedaje'] = df['answers'].apply(extraer_direccion)

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Confirmaciones", 
    "Campamento y Hospedaje", 
    "Alimentación", 
    "Mayoría de Edad", 
    "Representantes", 
    "Condiciones Médicas",
    "Exportar Excel"
])

with tab1:
    st.subheader("Confirmaciones por Sede")
    df_confirmados = df[df['confirm_submitted_at'].notna()]
    conf_sede = df_confirmados['university'].value_counts().reset_index()
    conf_sede.columns = ['Sede', 'Total Confirmados']
    st.dataframe(conf_sede, use_container_width=True)
    st.metric("Total Nacional Confirmados", len(df_confirmados))

with tab2:
    st.subheader("Censo de Campamento")
    camp_sede = pd.crosstab(df['university'], df['Campamento'], margins=True, margins_name='TOTAL NACIONAL')
    st.dataframe(camp_sede, use_container_width=True)
    
    st.subheader("Direcciones de No Acampantes")
    df_no_acampa = df[df['Campamento'] == 'No Acampa'][['first_name', 'last_name', 'university', 'phone', 'Direccion_Hospedaje']]
    st.dataframe(df_no_acampa, use_container_width=True)

with tab3:
    st.subheader("Tipos de Alimentación")
    alim_sede = pd.crosstab(df['university'], df['dietary_preference'], margins=True, margins_name='TOTAL NACIONAL')
    st.dataframe(alim_sede, use_container_width=True)

with tab4:
    st.subheader("Censo de Mayoría de Edad")
    edad_sede = pd.crosstab(df['university'], df['Es_Mayor'], margins=True, margins_name='TOTAL NACIONAL')
    st.dataframe(edad_sede, use_container_width=True)

with tab5:
    st.subheader("Censo de Representantes")
    df_repres = df[df['is_representative'] == True]
    rep_sede = df_repres['university'].value_counts().reset_index()
    rep_sede.columns = ['Sede', 'Total Representantes']
    st.dataframe(rep_sede, use_container_width=True)
    st.metric("Total Nacional Representantes", len(df_repres))

with tab6:
    st.subheader("Condiciones Médicas y Alergias")
    df_medico = df[(df['has_disability'] == True) | (df['allergies'] == True)]
    medico_sede = df_medico['university'].value_counts().reset_index()
    medico_sede.columns = ['Sede', 'Total Casos Médicos']
    st.dataframe(medico_sede, use_container_width=True)
    st.metric("Total Nacional Casos Médicos", len(df_medico))
    
    st.write("Detalle:")
    st.dataframe(df_medico[['first_name', 'last_name', 'university', 'has_disability', 'disability_specify', 'allergies', 'allergy_specify']], use_container_width=True)

with tab7:
    st.subheader("Generar Archivo Organizado")
    st.write("Presiona el botón para descargar todas estas tablas organizadas en un solo archivo de Excel.")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        conf_sede.to_excel(writer, sheet_name='Confirmaciones', index=False)
        camp_sede.to_excel(writer, sheet_name='Campamento')
        df_no_acampa.to_excel(writer, sheet_name='Direcciones', index=False)
        alim_sede.to_excel(writer, sheet_name='Alimentacion')
        edad_sede.to_excel(writer, sheet_name='Mayoria_Edad')
        rep_sede.to_excel(writer, sheet_name='Representantes', index=False)
        df_medico[['first_name', 'last_name', 'university', 'has_disability', 'disability_specify', 'allergies', 'allergy_specify']].to_excel(writer, sheet_name='Condiciones_Medicas', index=False)
    
    st.download_button(
        label="Descargar Reporte Excel",
        data=buffer.getvalue(),
        file_name="Reporte_Logistica_Completo.xlsx",
        mime="application/vnd.ms-excel"
    )
