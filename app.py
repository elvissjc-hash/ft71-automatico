import re
import sys
import pdfplumber
import pandas as pd

# Estratégias testadas em ordem: primeiro tabelas com bordas, depois tabelas sem bordas
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


def limpar(valor):
    """Remove None, quebras de linha e espaços duplicados."""
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def tabela_valida(linhas):
    """Descarta falsos positivos: precisa ter pelo menos 2 linhas e 2 colunas com conteúdo."""
    if len(linhas) < 2:
        return False
    if max(len(l) for l in linhas) < 2:
        return False
    return any(any(c and str(c).strip() for c in l) for l in linhas)


def remover_vazios(df):
    """Remove linhas e colunas totalmente vazias, sem alterar a ordem do conteúdo."""
    df = df.loc[~(df == "").all(axis=1)]
    df = df.loc[:, ~(df == "").all(axis=0)]
    return df.reset_index(drop=True)


def extrair_tabelas(caminho_pdf):
    """
    Lê todas as páginas do PDF e devolve uma lista de dicionários:
    {"pagina": int, "ordem": int, "df": DataFrame}
    """
    resultados = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for num_pag, pagina in enumerate(pdf.pages, start=1):
            tabelas = []
            for settings in ESTRATEGIAS:
                encontradas = pagina.extract_tables(table_settings=settings)
                tabelas = [t for t in encontradas if t and tabela_valida(t)]
                if tabelas:  # se achou com a primeira estratégia, não tenta a segunda
                    break

            for ordem, tabela in enumerate(tabelas, start=1):
                df = pd.DataFrame(tabela).map(limpar)
                df = remover_vazios(df)
                if not df.empty:
                    resultados.append({"pagina": num_pag, "ordem": ordem, "df": df})
    return resultados


def nome_aba(base, usados):
    """Gera nome válido para aba do Excel (até 31 caracteres, sem caracteres proibidos e sem duplicar)."""
    limpo = re.sub(r"[\\/*?:\[\]]", "", base)[:31]
    nome, i = limpo, 2
    while nome in usados:
        sufixo = f"_{i}"
        nome = limpo[: 31 - len(sufixo)] + sufixo
        i += 1
    usados.add(nome)
    return nome


def salvar_excel(resultados, caminho_saida):
    usados = set()
    indice = []

    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        for item in resultados:
            nome = nome_aba(f"P{item['pagina']}_T{item['ordem']}", usados)
            # header=False e index=False: a tabela sai exatamente como no PDF
            item["df"].to_excel(writer, sheet_name=nome, header=False, index=False)
            indice.append({
                "Aba": nome,
                "Página do PDF": item["pagina"],
                "Tabela na página": item["ordem"],
                "Linhas": item["df"].shape[0],
                "Colunas": item["df"].shape[1],
            })

        if not resultados:
            pd.DataFrame({"Aviso": ["Nenhuma tabela encontrada no PDF."]}).to_excel(
                writer, sheet_name="Aviso", index=False
            )
        else:
            pd.DataFrame(indice).to_excel(writer, sheet_name="Indice", index=False)

        # Ajusta largura das colunas em todas as abas
        for ws in writer.sheets.values():
            for coluna in ws.columns:
                tamanho = max((len(str(c.value)) for c in coluna if c.value is not None), default=8)
                ws.column_dimensions[coluna[0].column_letter].width = min(tamanho + 2, 60)

    return len(resultados)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python extrair_tabelas_pdf.py certificado.pdf [saida.xlsx]")
        sys.exit(1)

    pdf_entrada = sys.argv[1]
    excel_saida = sys.argv[2] if len(sys.argv) > 2 else "tabelas_certificado.xlsx"

    tabelas = extrair_tabelas(pdf_entrada)
    total = salvar_excel(tabelas, excel_saida)
    print(f"{total} tabela(s) encontrada(s). Arquivo salvo: {excel_saida}")
