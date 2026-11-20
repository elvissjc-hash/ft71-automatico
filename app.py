import streamlit as st

st.set_page_config(
    page_title="FT71 Automático",
    page_icon="📊",
    layout="wide"
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
    min_value=1,
    value=16
)

if st.button("PROCESSAR FT71"):

    if pdf_file is None:
        st.error("Selecione o PDF")
        st.stop()

    if ft71_file is None:
        st.error("Selecione a FT71")
        st.stop()

    st.success("Arquivos carregados")

    st.write("PDF:", pdf_file.name)
    st.write("FT71:", ft71_file.name)
    st.write("Registro:", registro)
    st.write("Linha Inicial:", linha_inicial)

    st.info(
        "Próxima etapa: integrar a lógica completa da FT71."
    )
