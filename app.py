import io
import re
import unicodedata

import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURAÇÃO
# ==========================================

st.set_page_config(
    page_title="Extrator de Resultados de Calibração",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Extrator de Resultados de Calibração")
st.markdown(
    "Copia apenas as tabelas de resultados de calibração do PDF e gera um Excel com elas uma embaixo da outra."
)

# Primeiro tenta tabelas com bordas; se não achar resultados, tenta alinhamento de texto
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

# Palavras que indicam cabeçalho de resultado de calibração (PT e EN), já sem acento e em minúsculas
PALAVRAS_RESULTADO = [
    "valor nominal", "nominal value",
    "erro", "error",
    "incerteza", "uncertainty",
    "ref",
]

# Títulos que indicam a seção de resultados na página
TITULOS_RESULTADO = [
    "resultados da calibracao",
    "calibration results",
]

# ==========================================
# FUNÇÕES
# ==========================================

def normalizar(txt):
    """Remove acentos, deixa em minúsculas e colapsa espaços."""
    if txt is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", txt).strip().lower()


def limpar_celula(valor):
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def tabela_valida(tabela):
    """Descarta falsos positivos: precisa ter ao menos 2 linhas com conteúdo e 2 colunas."""
    linhas = [l for l in tabela if any(limpar_celula(c) for c in l)]
    if len(linhas) < 2:
        return False
    return max(len(l) for l in linhas) >= 2


def e_resultado_calibracao(tabela, texto_pagina):
    """
    Considera a tabela um resultado de calibração se:
    - ao menos 2 palavras-chave de resultado aparecem nas células da tabela, ou
    - a página tem título de resultados e a tabela tem ao menos 2 colunas numéricas.
    """
    texto_tabela = normalizar(" ".join(limpar_celula(c) for linha in tabela for c in linha))
    acertos = sum(1 for p in PALAVRAS_RESULTADO if p in texto_tabela)
    if acertos >= 2:
        return True

    texto_pag = normalizar(texto_pagina)
    if any(t in texto_pag for t in TITULOS_RESULTADO):
        numericas = re.findall(r"\d+[.,]\d+", texto_tabela)
        return len(numericas) >= 4
    return False


def extrair_tabelas(arquivo_pdf, somente_calibracao=True):
    """
    Lê todas as páginas e devolve uma lista de dicionários:
    {"pagina": int, "ordem": int, "linhas": [[célula, ...], ...]}
    """
    resultados = []

    with pdfplumber.open(arquivo_pdf) as pdf:
        for num_pagina, pagina in enumerate(pdf.pages, start=1):
            texto_pagina = pagina.extract_text() or ""
            tabelas = []

            for settings in ESTRATEGIAS:
                encontradas = pagina.extract_tables(table_settings=settings)
                validas = [t for t in encontradas if t and tabela_valida(t)]

                if somente_calibracao:
                    validas = [t for t in validas if e_resultado_calibracao(t, texto_pagina)]

                if validas:
                    tabelas = validas
                    break

            for ordem, tabela in enumerate(tabelas, start=1):
                linhas = [[limpar_celula(c) for c in linha] for linha in tabela]
                linhas = [l for l in linhas if any(l)]

                if not linhas:
                    continue

                # Iguala o número de colunas
                n_colunas = max(len(l) for l in linhas)
                linhas = [l + [""] * (n_colunas - len(l)) for l in linhas]

                # Remove colunas totalmente vazias
                colunas_uteis = [i for i in range(n_colunas) if any(l[i] for l in linhas)]
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
    ws.title = "RESULTADOS"

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
        titulo = f"Página {item['pagina']} - Tabela {item['ordem']}"
        cel_titulo = ws.cell(row=linha_atual, column=1, value=titulo)
        cel_titulo.font = Font(bold=True)
        cel_titulo.fill = fill_titulo
        linha_atual += 1

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

somente_calibracao = st.checkbox(
    "Copiar somente tabelas de resultados de calibração",
    value=True
)

if st.button("🚀 EXTRAIR TABELAS"):

    if pdf_file is None:
        st.error("Selecione o PDF.")
        st.stop()

    with st.spinner("Lendo o PDF..."):
        resultados = extrair_tabelas(
            io.BytesIO(pdf_file.getvalue()),
            somente_calibracao=somente_calibracao
        )

    if not resultados:
        st.warning(
            "Nenhuma tabela de resultados de calibração foi encontrada. "
            "Desmarque a opção acima para ver todas as tabelas."
        )
        st.stop()

    st.success(f"{len(resultados)} tabela(s) copiada(s).")

    for item in resultados:
        st.subheader(f"Página {item['pagina']} - Tabela {item['ordem']}")
        st.dataframe(item["linhas"], use_container_width=True)

    excel = gerar_excel(resultados)

    st.download_button(
        label="📥 Baixar Excel",
        data=excel,
        file_name="RESULTADOS_CALIBRACAO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
