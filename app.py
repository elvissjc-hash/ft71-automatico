import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import io

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# --------------------------------------------------
# CONFIGURAÇÃO
# --------------------------------------------------

st.set_page_config(
    page_title="FT71 Automático",
    page_icon="📊",
    layout="wide"
)

st.title("📊 FT71 Automático")

st.markdown(
    "Preenchimento automático da FT71 a partir de certificado PDF"
)

# --------------------------------------------------
# FUNÇÕES
# --------------------------------------------------

def limpar(txt):

    import unicodedata

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


def detectar_grandeza(texto):

    txt = limpar(texto)

    regras = {

        "capacitance measurements":
        "Capacitância",

        "medicoes em capacitancia":
        "Capacitância",

        "resistance measurements":
        "Resistência",

        "medicoes em resistencia":
        "Resistência",

        "frequency measurements":
        "Frequência",

        "medicoes em frequencia":
        "Frequência",

        "voltage measurements":
        "Tensão",

        "medicoes em tensao":
        "Tensão",

        "current measurements":
        "Corrente",

        "medicoes em corrente":
        "Corrente",

        "inductance measurements":
        "Indutância"
    }

    for chave, valor in regras.items():

        if chave in txt:
            return valor

    return ""


def detectar_unidade(df):

    for coluna in df.columns:

        texto = str(coluna)

        if (
            "UUT" in texto
            or
            "Instrumento" in texto
        ):

            partes = texto.split("\n")

            for parte in partes:

                parte = parte.strip()

                if parte in [
                    "pF",
                    "nF",
                    "µF",
                    "uF",
                    "mF",
                    "F",
                    "Ω",
                    "kΩ",
                    "MΩ",
                    "V",
                    "mV",
                    "kV",
                    "A",
                    "mA"
                ]:

                    return parte

    return ""


def localizar_coluna(df, termos):

    for col in df.columns:

        nome = limpar(col)

        for termo in termos:

            if termo.lower() in nome:

                return col

    return None

# --------------------------------------------------
# INTERFACE
# --------------------------------------------------

pdf_file = st.file_uploader(
    "📄 Certificado PDF",
    type=["pdf"]
)

ft71_file = st.file_uploader(
    "📊 FT71",
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

# --------------------------------------------------
# PROCESSAR
# --------------------------------------------------

if st.button("🚀 PROCESSAR FT71"):

    if pdf_file is None:

        st.error(
            "Selecione o PDF"
        )

        st.stop()

    if ft71_file is None:

        st.error(
            "Selecione a FT71"
        )

        st.stop()

    barra = st.progress(0)

    # --------------------------
    # salvar temporários
    # --------------------------

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

    # --------------------------
    # extrair tabelas
    # --------------------------

    dados_tabelas = []

    with pdfplumber.open(pdf_path) as pdf:

        for pagina in pdf.pages:

            texto_pagina = (
                pagina.extract_text()
                or ""
            )

            tabs = pagina.extract_tables()

            if not tabs:
                continue

            for tabela in tabs:

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

    st.success(
        f"{len(dados_tabelas)} tabelas encontradas"
    )

    # --------------------------
    # abrir FT71
    # --------------------------

    wb = load_workbook(
        xlsm_path,
        keep_vba=True
    )

    ws = wb[registro]

    linha_atual = linha_inicial

    barra.progress(40)

    # --------------------------
    # cabeçalhos FT71
    # --------------------------

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

            cabecalhos[col] = limpar(
                valor
            )

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

        elif "instrumento" in nome:
            col_inst = c

        elif "referencia" in nome:
            col_ref = c

        elif "incerteza" in nome:
            col_inc = c

        elif "erro" in nome:
            col_erro = c

    barra.progress(50)

    # --------------------------
    # preenchimento
    # --------------------------

    for bloco in dados_tabelas:

        texto_pagina = bloco["texto"]

        tabela = bloco["tabela"]

        grandeza = detectar_grandeza(
            texto_pagina
        )

        unidade = detectar_unidade(
            tabela
        )

        col_pdf_inst = localizar_coluna(
            tabela,
            ["uut", "instrumento"]
        )

        col_pdf_ref = localizar_coluna(
            tabela,
            ["ref", "reference"]
        )

        col_pdf_erro = localizar_coluna(
            tabela,
            ["erro", "error"]
        )

        col_pdf_inc = localizar_coluna(
            tabela,
            [
                "incerteza",
                "uncertainty"
            ]
        )

        for _, linha in tabela.iterrows():

            if col_grandeza:

                cel = ws.cell(
                    linha_atual,
                    col_grandeza
                )

                if not isinstance(
                    cel,
                    MergedCell
                ):

                    cel.value = grandeza

            if col_unidade:

                ws.cell(
                    linha_atual,
                    col_unidade,
                    unidade
                )

            if (
                col_inst
                and
                col_pdf_inst
            ):

                ws.cell(
                    linha_atual,
                    col_inst,
                    linha[col_pdf_inst]
                )

            if (
                col_ref
                and
                col_pdf_ref
            ):

                ws.cell(
                    linha_atual,
                    col_ref,
                    linha[col_pdf_ref]
                )

            if (
                col_inc
                and
                col_pdf_inc
            ):

                ws.cell(
                    linha_atual,
                    col_inc,
                    linha[col_pdf_inc]
                )

            if (
                col_erro
                and
                col_pdf_erro
            ):

                ws.cell(
                    linha_atual,
                    col_erro,
                    linha[col_pdf_erro]
                )

            linha_atual += 1

    barra.progress(90)

    # --------------------------
    # gerar download
    # --------------------------

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    barra.progress(100)

    st.success(
        "FT71 preenchida com sucesso"
    )

    st.download_button(
        label="📥 Baixar FT71_PREENCHIDA.xlsm",
        data=output,
        file_name="FT71_PREENCHIDA.xlsm",
        mime="application/vnd.ms-excel.sheet.macroEnabled.12"
    )
