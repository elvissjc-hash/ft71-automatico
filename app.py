import streamlit as st

st.set_page_config(
    page_title="FT71 Automático",
    page_icon="📊"
)

st.title("FT71 Automático")

pdf_file = st.file_uploader(
    "Selecione o Certificado PDF",
    type=["pdf"]
)

ft71_file = st.file_uploader(
    "Selecione a FT71",
    type=["xlsm"]
)

registro = st.selectbox(
    "Registro",
    [
        "REGISTRO_Ano_X",
        "REGISTRO_Ano_Y",
        "REGISTRO_Ano_Z"
    ]
)

linha_inicial = st.number_input(
    "Linha Inicial",
    value=16
)

if st.button("PROCESSAR FT71"):

    st.success("Integração iniciada!")

    st.write("PDF:", pdf_file)
    st.write("FT71:", ft71_file)
    st.write("Registro:", registro)
    st.write("Linha:", linha_inicial)
