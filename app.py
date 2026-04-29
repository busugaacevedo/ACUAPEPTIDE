import streamlit as st
import pandas as pd
from io import BytesIO
# Importamos tu motor de Word
from ACUAPEPTIDE_code import create_word
st.set_page_config(page_title="ACUAPEPTIDE")

st.title("🧬 ACUAPEPTIDE SÍNTESIS ✍️")
st.caption("***Version 2026. Creada por: Brandon Usuga-Acevedo***👨🏾‍🔬")

# --- SECCIÓN 1: INFORMACIÓN GENERAL ---
with st.expander("📝 Información del Proyecto", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        nameProject = st.text_input("Nombre del Proyecto", value="S220426")
        nameResin = st.text_input("Nombre de la Resina", value="RinkAmide")
    with col2:
        massResin = st.number_input("Masa de Resina (mg)", value=40)
        StResin = st.number_input("Sustitución (mmol/g)", value=0.67, format="%.2f")

# --- SECCIÓN 2: CONFIGURACIÓN DE QUÍMICA ---
with st.expander("🧪 Configuración de Reactivos"):
    deprotection = st.text_input("Desprotección", "Piperidina 20% TritonX100 1% en DMF")
    simple = st.text_input("Activador Simple", "AA + TBTU + OXYMA + DIEA")
    doble = st.text_input("Activador Doble", "AA + HBTU + OXYMA + DIEA")
    triple = st.text_input("Activador Triple", "AA + HCTU(DIC) + OXYMA + DIEA")

# --- SECCIÓN 3: CARGA DE DATOS ---
st.subheader("📂 Cargar Secuencias")
st.info("""
El archivo Excel (.xlsx) debe contener exactamente estas columnas:

\t Bolsas  |  Secuencia  |  Familia\n""")

uploaded_file = st.file_uploader("Sube tu Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    columnas_requeridas = ["Bolsas", "Secuencia", "Familia"]
    faltantes = [col for col in columnas_requeridas if col not in df.columns]
    if faltantes:
            st.error(f"❌ El archivo no contiene las columnas obligatorias: {', '.join(faltantes)}")
            st.stop()
    st.success("✅ Archivo cargado correctamente\n")

    st.write("Vista previa de péptidos:")
    st.dataframe(df.head(),use_container_width=True)
#    st.write("***HASTA ACÁ TODO VA BIEN...Linea37***")
    
    if st.button("🚀 Generar Documento de Síntesis"):
        try:
            # Extraer listas necesarias para tu función create_word
            bolsas_list = df["Bolsas"].tolist()
            peptides_list = df["Secuencia"].tolist()
            family_list = df["Familia"].tolist()
            
            # Llamar a tu función modificada
            # Nota: le pasamos un nombre de archivo temporal para el metadato
            output_buffer = create_word(nameProject, deprotection, nameResin, massResin, StResin, bolsas_list, peptides_list, family_list, f"{nameProject}.docx", simple, doble, triple )
            
            st.success("¡Documento generado con éxito!  🍻")
            
            st.download_button(
                label="📥 Descargar Documento de Síntesis (Word)",
                data=output_buffer,
                file_name=f"doc_{nameProject}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error al procesar: {e}")
            st.info("Asegúrate de que el Excel tenga las columnas 'Secuencia' y 'Familia'")
