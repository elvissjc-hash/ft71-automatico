import streamlit as st
#teste

#teste

import pdfplumber
import pandas as pd
import tempfile
import io
import unicodedata

from openpyxl import Workbook

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

    texto = limpar(texto)

    cabecalhos = ""

    if tabela is not None:

        cabecalhos = " ".join(
            [str(c) for c in tabela.columns]
        )

        cabecalhos = limpar(cabecalhos)

    conteudo = texto + " " + cabecalhos

    if "resistencia" in conteudo:
        return "RESISTENCIA"

    if (
        "corrente continua" in conteudo
        or
        "dc current" in conteudo
    ):
        return "CORRENTE DC"

    if (
        "corrente alternada" in conteudo
        or
        "ac current" in conteudo
    ):
        return "CORRENTE AC"

    if (
        "tensao continua" in conteudo
        or
        "tensao corrente continua" in conteudo
        or
        "dc voltage" in conteudo
    ):
        return "TENSAO DC"

    if (
        "tensao alternada" in conteudo
        or
        "tensao corrente alternada" in conteudo
        or
        "ac voltage" in conteudo
    ):
        return "TENSAO AC"

    if "capacitancia" in conteudo:
        return "CAPACITANCIA"

    if "frequencia" in conteudo:
        return "FREQUENCIA"

    if "medicoes lineares" in conteudo:
        return "DIMENSIONAL"

    return ""

def detectar_unidade(grandeza):

    mapa = {

        "RESISTENCIA": "Ω",

        "CAPACITANCIA": "pF",

        "FREQUENCIA": "Hz",

        "CORRENTE DC": "A",
        "CORRENTE AC": "A",

        "TENSAO DC": "V",
        "TENSAO AC": "V",

        "DIMENSIONAL": "mm"
    }

    return mapa.get(
        grandeza,
        ""
    )


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



# ==========================================
# PROCESSAMENTO
# ==========================================

if st.button("🚀 PROCESSAR FT71"):

    if pdf_file is None:
        st.error("Selecione o PDF")
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

    wb_saida = Workbook()

    ws_saida = wb_saida.active

    ws_saida.title = "IMPORTACAO"

    ws_saida["A1"] = "GRANDEZA"
    ws_saida["B1"] = "UNIDADE"
    ws_saida["C1"] = "INSTRUMENTO"
    ws_saida["D1"] = "REFERENCIA"
    ws_saida["E1"] = "ERRO"
    ws_saida["F1"] = "INCERTEZA"

    # informações para a macro VBA

    
    linha_saida = 2

    barra.progress(50)

    for bloco in dados_tabelas:

        texto_pagina = bloco["texto"]

        tabela = bloco["tabela"]

        grandeza = detectar_grandeza(
            texto_pagina,
            tabela
        )

        unidade = detectar_unidade(
            grandeza
        )

        col_pdf_inst = localizar_coluna(
            tabela,
            [
                    "uut",
                    "instrumento",
                    "valor nominal",
                    "valor do instrumento",
                    "valor indicado",
                    "reading",
                    "indication"
            ]
        )

        col_pdf_ref = localizar_coluna(
            tabela,
            [
                "ref",
                "reference",
                "referencia",
                "valor de referencia",
                "standard"
            ]
        )

        col_pdf_erro = localizar_coluna(
            tabela,
            [
                "erro",
                "error",
                "desvio",
                "deviation",
                "correcao"
            ]
        )

        col_pdf_inc = localizar_coluna(
            tabela,
            [
                "incerteza",
                "uncertainty",
                "expanded uncertainty"
            ]
        )

        for _, linha in tabela.iterrows():

            ws_saida.cell(
                linha_saida,
                1,
                grandeza
            )

            ws_saida.cell(
                linha_saida,
                2,
                unidade
            )

            if col_pdf_inst is not None:

                ws_saida.cell(
                    linha_saida,
                    3,
                    linha[col_pdf_inst]
                )

            if col_pdf_ref is not None:

                ws_saida.cell(
                    linha_saida,
                    4,
                    linha[col_pdf_ref]
                )

            if col_pdf_erro is not None:

                ws_saida.cell(
                    linha_saida,
                    5,
                    linha[col_pdf_erro]
                )

            if col_pdf_inc is not None:

                ws_saida.cell(
                    linha_saida,
                    6,
                    linha[col_pdf_inc]
                )

            linha_saida += 1

    barra.progress(90)

    output = io.BytesIO()

    wb_saida.save(output)

    output.seek(0)

    barra.progress(100)

    st.success(
        "Arquivo de importação criado com sucesso."
    )

    st.download_button(
        label="📥 Baixar DADOS_IMPORTACAO.xlsx",
        data=output,
        file_name="DADOS_IMPORTACAO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
