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

    txt = limpar(texto)

    if "medidor de resistencia" in txt:
        return "RESISTENCIA"

    if "medidor de corrente continua" in txt:
        return "CORRENTE DC"

    if "medidor de corrente alternada" in txt:
        return "CORRENTE AC"

    if "medidor de tensao corrente continua" in txt:
        return "TENSAO DC"

    if "medidor de tensao corrente alternada" in txt:
        return "TENSAO AC"

    if "capacitancia" in txt:
        return "CAPACITANCIA"

    if "frequencia" in txt:
        return "FREQUENCIA"

    if "medicoes lineares" in txt:
        return "DIMENSIONAL"

    linhas = texto.split("\n")

    for i, linha in enumerate(linhas):

        if "instrumento" in linha.lower():

            if i + 1 < len(linhas):

                proxima = linhas[i + 1].strip()

                if len(proxima) > 3:

                    return proxima

    return ""

def detectar_unidade(texto):

    txt = limpar(texto)

    if "resistencia" in txt:
        return "Ω"

    if "capacitancia" in txt:
        return "pF"

    if "frequencia" in txt:
        return "Hz"

    if "corrente continua" in txt:
        return "A"

    if "corrente alternada" in txt:
        return "A"

    if "tensao corrente continua" in txt:
        return "V"

    if "tensao corrente alternada" in txt:
        return "V"

    if "medicoes lineares" in txt:
        return "mm"

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
            texto_pagina
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
