import streamlit as st

import pdfplumber
import pandas as pd
import tempfile
import io
import re
import unicodedata

from openpyxl import Workbook


# ====================================================
# CONFIG
# ====================================================

st.set_page_config(
    page_title="FT71 Automático",
    page_icon="📊",
    layout="wide"
)

st.title("📊 FT71 Automático")
st.write("Importação automática de certificados RBC")


# ====================================================
# FUNÇÕES AUXILIARES
# ====================================================

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


# ====================================================
# MAPAS DE COLUNAS
# ====================================================

MAPA_INSTRUMENTO = [

    "instrumento",
    "uut",
    "valor nominal",
    "valor do instrumento",
    "valor indicado",
    "indication",
    "reading",
    "medido"
]

MAPA_REFERENCIA = [

    "ref",
    "referencia",
    "referência",
    "reference",
    "valor de referencia",
    "valor de referência",
    "valor padrao",
    "valor padrão",
    "standard"
]

MAPA_ERRO = [

    "erro",
    "error",
    "desvio",
    "deviation",
    "correcao",
    "correção",
    "correction",
]

MAPA_INCERTEZA = [

    "incerteza",
    "uncertainty",
    "expanded uncertainty",
    "u(k=2)",
    "u95"
]


# ====================================================
# LOCALIZAR COLUNAS
# ====================================================

def localizar_coluna(df, lista):

    for col in df.columns:

        nome = limpar(col)

        for termo in lista:

            if limpar(termo) in nome:
                return col

    return None


# ====================================================
# GRANDEZA
# ====================================================

def detectar_grandeza(texto):

    texto = limpar(texto)

    mapa = {

        "resistencia": "RESISTENCIA",
        "resistência": "RESISTENCIA",

        "capacitancia": "CAPACITANCIA",
        "capacitância": "CAPACITANCIA",

        "indutancia": "INDUTANCIA",

        "corrente continua": "CORRENTE DC",
        "corrente dc": "CORRENTE DC",

        "corrente alternada": "CORRENTE AC",
        "corrente ac": "CORRENTE AC",

        "tensao dc": "TENSAO DC",
        "tensao continua": "TENSAO DC",

        "tensao ac": "TENSAO AC",
        "tensao alternada": "TENSAO AC",

        "frequencia": "FREQUENCIA",

        "medicoes lineares": "DIMENSIONAL",
        "medicao de pecas": "DIMENSIONAL",
        "medição de peças": "DIMENSIONAL"
    }

    for chave, valor in mapa.items():

        if chave in texto:
            return valor

    return "NAO IDENTIFICADA"


# ====================================================
# UNIDADE
# ====================================================

def detectar_unidade(texto):

    unidades = [

        "mm",
        "cm",
        "m",

        "um",
        "µm",

        "pf",
        "nf",
        "uf",
        "µf",

        "ohm",
        "ω",
        "kω",
        "mω",

        "v",
        "mv",
        "kv",

        "a",
        "ma",
        "ua",

        "hz",
        "khz",
        "mhz",
        "ghz"
    ]

    txt = limpar(texto)

    for u in unidades:

        if u in txt:
            return u

    return ""



# ====================================================
# FALLBACK PDFPLUMBER
# ====================================================

def extrair_pdfplumber(arquivo_pdf):

    tabelas_extraidas = []

    with pdfplumber.open(arquivo_pdf) as pdf:

        for pagina in pdf.pages:

            tabelas = pagina.extract_tables()

            for tb in tabelas:

                try:

                    if len(tb) > 1:

                        df = pd.DataFrame(
                            tb[1:],
                            columns=tb[0]
                        )

                        tabelas_extraidas.append(df)

                except:
                    pass

    return tabelas_extraidas


# ====================================================
# UPLOAD
# ====================================================

pdf_file = st.file_uploader(
    "Selecione o PDF",
    type=["pdf"]
)

# ====================================================
# PROCESSAMENTO
# ====================================================

if st.button("🚀 PROCESSAR FT71"):

    if pdf_file is None:

        st.error("Selecione um PDF")

        st.stop()

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".pdf"
    ) as tmp:

        tmp.write(pdf_file.getvalue())

        caminho_pdf = tmp.name

    with pdfplumber.open(caminho_pdf) as pdf:

        paginas_texto = []

        for p in pdf.pages:

            paginas_texto.append(
                 p.extract_text() or ""
            )

    tabelas = extrair_pdfplumber(
    caminho_pdf
    )

    def detectar_grandeza_pagina(texto):

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

    return "NAO IDENTIFICADA"


    def detectar_unidade_grandeza(grandeza):

        unidades = {

        "RESISTENCIA": "Ω",

        "CAPACITANCIA": "pF",

        "FREQUENCIA": "Hz",

        "TENSAO DC": "V",
        "TENSAO AC": "V",

        "CORRENTE DC": "A",
        "CORRENTE AC": "A",

        "DIMENSIONAL": "mm"
    }

    return unidades.get(
        grandeza,
        ""
    )

    wb = Workbook()

    ws = wb.active

    ws.title = "IMPORTACAO"

    ws.append(

        [
            "GRANDEZA",
            "UNIDADE",
            "INSTRUMENTO",
            "REFERENCIA",
            "ERRO",
            "INCERTEZA"
        ]
    )

    registros = 0
    
    indice_pagina = 0
    for tabela in tabelas:

    pagina_texto = ""

    if indice_pagina < len(paginas_texto):

        pagina_texto = paginas_texto[
            indice_pagina
        ]

    grandeza = detectar_grandeza_pagina(
        pagina_texto
    )

    unidade = detectar_unidade_grandeza(
        grandeza
    )

    indice_pagina += 1

        col_inst = localizar_coluna(
            tabela,
            MAPA_INSTRUMENTO
        )

        col_ref = localizar_coluna(
            tabela,
            MAPA_REFERENCIA
        )

        col_erro = localizar_coluna(
            tabela,
            MAPA_ERRO
        )

        col_inc = localizar_coluna(
            tabela,
            MAPA_INCERTEZA
        )

        if (
            col_inst is None
            and col_ref is None
            and col_erro is None
        ):
            continue

        for _, row in tabela.iterrows():

            ws.append(

                [
                    grandeza,
                    unidade,

                    row[col_inst]
                    if col_inst else "",

                    row[col_ref]
                    if col_ref else "",

                    row[col_erro]
                    if col_erro else "",

                    row[col_inc]
                    if col_inc else ""
                ]
            )

            registros += 1

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    st.success(
        f"{registros} registros importados."
    )

    st.download_button(
        label="📥 Baixar DADOS_IMPORTACAO.xlsx",
        data=output,
        file_name="DADOS_IMPORTACAO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
