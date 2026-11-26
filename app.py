import io
import re

import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURAÇÃO
# ==========================================

st.set_page_config(
    page_title="Extrator de Tabelas",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Extrator de Tabelas de Certificados")
st.markdown(
    "Extrai todas as tabelas do PDF e gera um Excel com elas uma embaixo da outra."
)

# Primeiro tenta tabelas com bordas; se não achar, tenta tabelas só com alinhamento de texto
ESTRATEGIAS = [
    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
    {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "intersection_tolerance": 5,
    },
]

# ==========================================
# FUNÇÕES
# ==========================================

def limpar_celula(valor):
    """Converte para texto, remove quebras de linha e espaços duplicados."""
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def tabela_valida(tabela):
    """Descarta falsos positivos: precisa ter ao menos 2 linhas com conteúdo e 2 colunas."""
    linhas = [l for l in tabela if any(limpar_celula(c) for c in l)]
    if len(linhas) < 2:
        return False
    return max(len(l) for l in linhas) >= 2


def extrair_tabelas(arquivo_pdf):
    """
    Lê todas as páginas e devolve uma lista de dicionários:
    {"pagina": int, "ordem": int, "linhas": [[célula, ...], ...]}
    """
    resultados = []

    with pdfplumber.open(arquivo_pdf) as pdf:
        for num_pagina, pagina in enumerate(pdf.pages, start=1):
            tabelas = []

            for settings in ESTRATEGIAS:
                encontradas = pagina.extract_tables(table_settings=settings)
                tabelas = [t for t in encontradas if t and tabela_valida(t)]
                if tabelas:
                    break

            for ordem, tabela in enumerate(tabelas, start=1):
                linhas = [[limpar_celula(c) for c in linha] for linha in tabela]
                linhas = [l for l in linhas if any(l)]  # remove linhas vazias

                if not linhas:
                    continue

                # Iguala o número de colunas das linhas
                n_colunas = max(len(l) for l in linhas)
                linhas = [l + [""] * (n_colunas - len(l)) for l in linhas]

                # Remove colunas totalmente vazias
                colunas_uteis = [
                    i for i in range(n_colunas)
                    if any(l[i] for l in linhas)
                ]
                linhas = [[l[i] for i in colunas_uteis] for l in linhas]

                resultados.append({
                    "pagina": num_pagina,
                    "ordem": ordem,
                    "linhas": linhas,
                })

    return resultados


def gerar_excel(resultados):
    """Grava todas as tabelas em uma aba, uma embaixo da outra."""
    wb = Workbook()
    ws = wb.active
    ws.title = "TABELAS"

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    fill_titulo = PatternFill("solid", fgColor="BDD7EE")
    largura = {}
    linha_atual = 1

    for item in resultados:
        # Título da tabela
        titulo = f"Página {item['pagina']} - Tabela {item['ordem']}"
        cel_titulo = ws.cell(row=linha_atual, column=1, value=titulo)
        cel_titulo.font = Font(bold=True)
        cel_titulo.fill = fill_titulo
        linha_atual += 1

        # Linhas da tabela, exatamente como no PDF
        for registro in item["linhas"]:
            for c, valor in enumerate(registro, start=1):
                cel = ws.cell(row=linha_atual, column=c, value=valor)
                cel.border = borda
                cel.alignment = Alignment(vertical="center", wrap_text=True)
                largura[c] = max(largura.get(c, 0), len(valor))
            linha_atual += 1

        linha_atual += 1  # linha em branco entre tabelas

    for c, tamanho in largura.items():
        ws.column_dimensions[get_column_letter(c)].width = min(max(tamanho + 2, 10), 60)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==========================================
# INTERFACE
# ==========================================

pdf_file = st.file_uploader("📄 Certificado PDF", type=["pdf"])

if st.button("🚀 EXTRAIR TABELAS"):

    if pdf_file is None:
        st.error("Selecione o PDF.")
        st.stop()

    with st.spinner("Lendo o PDF..."):
        resultados = extrair_tabelas(io.BytesIO(pdf_file.getvalue()))

    if not resultados:
        st.warning("Nenhuma tabela foi encontrada neste PDF.")
        st.stop()

    st.success(f"{len(resultados)} tabela(s) encontrada(s).")

    # Pré-visualização de cada tabela
    for item in resultados:
        st.subheader(f"Página {item['pagina']} - Tabela {item['ordem']}")
        st.dataframe(item["linhas"], use_container_width=True)

    excel = gerar_excel(resultados)

    st.download_button(
        label="📥 Baixar Excel com todas as tabelas",
        data=excel,
        file_name="TABELAS_CERTIFICADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
