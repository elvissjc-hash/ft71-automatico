import streamlit as st
#teste
import platform
import streamlit as st

st.write(platform.system())
st.write(platform.platform()) 


#teste

import pdfplumber
import pandas as pd
import tempfile
import io
import unicodedata

from openpyxl import load_workbook

# ==========================================
# CONFIGURAÇÃO
# ==========================================

st.set_page_config(
    page_title="FT71 Automático",
    page_icon="📊",
    layout="wide"
)

st.title("📊 FT71 Automático")

st.markdown(
    "Preenchimento automático da FT71 a partir de certificados PDF"
)

# ==========================================
# FUNÇÕES
# ==========================================

def limpar(txt):

    if txt is None:
        return ""

    txt = str(txt)

    txt = unicodedata.normalize(
        "NFKD",
        txt
    ).encode(
        "ascii",
        "ignore"
    ).decode()

    txt = txt.lower()

    txt = txt.replace("\n", " ")

    return txt.strip()


def detectar_grandeza(texto, tabela=None):

    linhas = texto.split("\n")

    # prioridade 1

    for i, linha in enumerate(linhas):

        if "instrumento" in linha.lower():

            if i + 1 < len(linhas):

                proxima = linhas[i + 1].strip()

                if len(proxima) > 3:

                    return proxima

    # prioridade 2

    for linha in linhas:

        linha_limpa = linha.strip()

        if "medições" in linha_limpa.lower():
            return linha_limpa

        if "measurement" in linha_limpa.lower():
            return linha_limpa

    return ""


def detectar_unidade(texto):

    unidades = [
        "mm",
        "cm",
        "m",
        "µm",
        "pF",
        "nF",
        "µF",
        "uF",
        "mF",
        "F",
        "Hz",
        "kHz",
        "MHz",
        "GHz",
        "V",
        "mV",
        "kV",
        "A",
        "mA",
        "Ω",
        "kΩ",
        "MΩ"
    ]

    texto = str(texto)

    for unidade in unidades:

        if unidade in texto:
            return unidade

    return ""


def localizar_coluna(df, termos):

    for col in df.columns:

        nome = limpar(col)

        for termo in termos:

            if termo.lower() in nome:

                return col

    return None


# ==========================================
# INTERFACE
# ==========================================

pdf_file = st.file_uploader(
    "📄 Certificado PDF",
    type=["pdf"]
)

ft71_file = st.file_uploader(
    "📊 FT71 (.xlsm)",
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

linha_atual = st.number_input(
    "Linha Inicial",
    value=16,
    min_value=1
)

# ==========================================
# PROCESSAMENTO
# ==========================================

if st.button("🚀 PROCESSAR FT71"):

    if pdf_file is None:

        st.error("Selecione o PDF")
        st.stop()

    if ft71_file is None:

        st.error("Selecione a FT71")
        st.stop()

    barra = st.progress(0)

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".pdf"
    ) as tmp_pdf:

        tmp_pdf.write(
            pdf_file.getvalue()
        )

        pdf_path = tmp_pdf.name

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsm"
    ) as tmp_xlsm:

        tmp_xlsm.write(
            ft71_file.getvalue()
        )

        xlsm_path = tmp_xlsm.name

    barra.progress(10)

    dados_tabelas = []

    with pdfplumber.open(pdf_path) as pdf:

        for pagina in pdf.pages:

            texto_pagina = (
                pagina.extract_text()
                or ""
            )

            tabelas = pagina.extract_tables()

            if not tabelas:
                continue

            for tabela in tabelas:

                try:

                    df = pd.DataFrame(
                        tabela[1:],
                        columns=tabela[0]
                    )

                    if len(df) > 0:

                        dados_tabelas.append(
                            {
                                "texto": texto_pagina,
                                "tabela": df
                            }
                        )

                except:
                    pass

    barra.progress(30)

    wb = load_workbook(
        xlsm_path,
        keep_vba=True
    )

    ws = wb[registro]

    cabecalhos = {}

    for col in range(
        1,
        ws.max_column + 1
    ):

        valor = ws.cell(
            row=15,
            column=col
        ).value

        if valor:

            cabecalhos[col] = limpar(valor)

    col_grandeza = None
    col_unidade = None
    col_inst = None
    col_ref = None
    col_inc = None
    col_erro = None

    for c, nome in cabecalhos.items():

        if "grandeza" in nome:
            col_grandeza = c

        elif "unidade" in nome:
            col_unidade = c

        elif (
            "instrumento" in nome
            or
            "valor do instrumento" in nome
        ):
            col_inst = c

        elif "referencia" in nome:
            col_ref = c

        elif "incerteza" in nome:
            col_inc = c

        elif "erro" in nome:
            col_erro = c

    barra.progress(50)

    for bloco in dados_tabelas:

        texto_pagina = bloco["texto"]

        tabela = bloco["tabela"]

        grandeza = detectar_grandeza(
            texto_pagina,
            tabela
        )

        unidade = detectar_unidade(
            texto_pagina
        )

        col_pdf_inst = localizar_coluna(
            tabela,
            [
                "uut",
                "instrumento",
                "valor nominal"
            ]
        )

        col_pdf_ref = localizar_coluna(
            tabela,
            [
                "ref",
                "reference"
            ]
        )

        col_pdf_erro = localizar_coluna(
            tabela,
            [
                "erro",
                "error"
            ]
        )

        col_pdf_inc = localizar_coluna(
            tabela,
            [
                "incerteza",
                "uncertainty"
            ]
        )

        for _, linha in tabela.iterrows():

            if col_grandeza is not None:

                ws.cell(
                    row=linha_atual,
                    column=col_grandeza,
                    value=grandeza
                )

            if col_unidade is not None:

                ws.cell(
                    row=linha_atual,
                    column=col_unidade,
                    value=unidade
                )

            if (
                col_inst is not None
                and
                col_pdf_inst is not None
            ):

                ws.cell(
                    row=linha_atual,
                    column=col_inst,
                    value=linha[col_pdf_inst]
                )

            if (
                col_ref is not None
                and
                col_pdf_ref is not None
            ):

                ws.cell(
                    row=linha_atual,
                    column=col_ref,
                    value=linha[col_pdf_ref]
                )

            if (
                col_inc is not None
                and
                col_pdf_inc is not None
            ):

                ws.cell(
                    row=linha_atual,
                    column=col_inc,
                    value=linha[col_pdf_inc]
                )

            if (
                col_erro is not None
                and
                col_pdf_erro is not None
            ):

                ws.cell(
                    row=linha_atual,
                    column=col_erro,
                    value=linha[col_pdf_erro]
                )

            linha_atual += 1

     barra.progress(90)

     output = io.BytesIO()

     wb.save(output)
   
     output.seek(0)
    
    # wb.save(xlsm_path)

barra.progress(100)

st.success(
    "FT71 preenchida com sucesso."
)

st.download_button(
    label="📥 Baixar FT71_PREENCHIDA.xlsm",
    data=output,
    file_name="FT71_PREENCHIDA.xlsm",
    mime="application/vnd.ms-excel.sheet.macroEnabled.12"
)
