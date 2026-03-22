import streamlit as st
import pandas as pd
import json
import io

st.set_page_config(layout="wide")

try:
    df = pd.read_excel('Datos-Sabado-Noche.xlsx')
    df.columns = df.columns.str.strip()
except Exception as e:
    st.error(f"Error al leer el archivo: {e}")
    st.stop()

if 'confirm_submitted_at' not in df.columns:
    st.error("Error: No se encontró la columna 'confirm_submitted_at'.")
    st.stop()

df_conf = df[df['confirm_submitted_at'].notna()].copy()

df_conf['Es_Mayor'] = df_conf['document_type'].apply(lambda x: 'Mayor' if str(x).strip().upper() == 'CC' else 'Menor')

def check_campamento(ans):
    texto = str(ans)
    if 'Planea acampar' in texto:
        return 'Acampa'
    return 'No Acampa'

def get_direccion(ans):
    try:
        d = json.loads(ans)
        for key, value in d.items():
            if 'direccion' in key.lower() or 'address' in key.lower():
                return value
        return 'No especifica'
    except:
        return 'Sin datos'

df_conf['Campamento'] = df_conf['confirm_answers'].apply(check_campamento)
df_conf['Direccion'] = df_conf['confirm_answers'].apply(get_direccion)

def es_afirmativo(valor):
    v = str(valor).lower().strip()
    return v in ['si', 'sí', 'true']

def check_representante(row):
    cuerpo = str(row.get('representative_bodies', '')).lower()
    if 'consejo' in cuerpo:
        return 'Sí'
    return 'No'

df_conf['Repre_Str'] = df_conf.apply(check_representante, axis=1)

df_conf['Tiene_Alergia'] = df_conf['allergies'].apply(lambda x: es_afirmativo(x))

col1, col2, col3 = st.columns(3)

with col1:
    with st.container(border=True):
        st.subheader("Confirmaciones")
        conf_sede = df_conf['university'].value_counts().reset_index()
        conf_sede.columns = ['Sede', 'Total']
        st.dataframe(conf_sede, hide_index=True, width='stretch')
        st.write(f"Total Nacional: {len(df_conf)}")

with col2:
    with st.container(border=True):
        st.subheader("Mayoría de Edad")
        edad_sede = pd.crosstab(df_conf['university'], df_conf['Es_Mayor'], margins=True, margins_name='Nacional')
        st.dataframe(edad_sede, width='stretch')

with col3:
    with st.container(border=True):
        st.subheader("Condiciones Médicas (Alergias)")
        df_med = df_conf[df_conf['Tiene_Alergia'] == True]
        if len(df_med) > 0:
            med_sede = df_med['university'].value_counts().reset_index()
            med_sede.columns = ['Sede', 'Total Casos']
            st.dataframe(med_sede, hide_index=True, width='stretch')
        else:
            st.write("No hay registros de alergias.")
        st.write(f"Total Nacional: {len(df_med)}")

col4, col5, col6 = st.columns(3)

with col4:
    with st.container(border=True):
        st.subheader("Campamento")
        camp_sede = pd.crosstab(df_conf['university'], df_conf['Campamento'], margins=True, margins_name='Nacional')
        st.dataframe(camp_sede, width='stretch')
        
        st.write("Muestra de No Acampantes:")
        df_no_acampa = df_conf[df_conf['Campamento'] == 'No Acampa'][['first_name', 'university', 'Direccion']].head(5)
        st.dataframe(df_no_acampa, hide_index=True, width='stretch')

with col5:
    with st.container(border=True):
        st.subheader("Alimentación")
        alim_sede = pd.crosstab(df_conf['university'], df_conf['dietary_preference'], margins=True, margins_name='Nacional')
        st.dataframe(alim_sede, width='stretch')

with col6:
    with st.container(border=True):
        st.subheader("Representantes")
        rep_sede = pd.crosstab(df_conf['university'], df_conf['Repre_Str'], margins=True, margins_name='Nacional')
        st.dataframe(rep_sede, width='stretch')

st.write("")

with st.container(border=True):
    col_txt, col_btn = st.columns([3, 1])
    
    with col_txt:
        st.write("Descargar toda la base de datos de confirmación agregando si son repres o no, tipo de documento, correo, dieta, y respuestas de confirmación.")
    
    with col_btn:
        columnas_descarga = [
            'first_name', 'last_name', 'document_type', 'document_number', 
            'email', 'phone', 'university', 'dietary_preference', 
            'Repre_Str', 'representative_bodies', 'Campamento', 'Direccion', 'allergies', 'confirm_answers'
        ]
        
        df_descarga = df_conf[[c for c in columnas_descarga if c in df_conf.columns]]
        
        buffer = io.BytesIO()
        df_descarga.to_excel(buffer, index=False, engine='xlsxwriter')
        st.download_button(
            label="Descargar Base Logística",
            data=buffer.getvalue(),
            file_name="Logistica_ENEUN_Organizado.xlsx",
            mime="application/vnd.ms-excel",
            use_container_width=True 
        )
