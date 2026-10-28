from datetime import datetime, timedelta
import json
import logging
from pathlib import Path

from typing import Dict, Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd

from relat import ExcelReport
import ajusta_painel_sla as aps
import gerar_total_sla as gts

# ==============================================================================
# CONFIGURAÇÃO DE AMBIENTE E LOGGING
# ==============================================================================
PASTA_TRABALHO = Path(__file__).resolve().parent
logger = logging.getLogger("SLA")
logger.setLevel(logging.INFO)

formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

# Handler de arquivo
arquivo_log = logging.FileHandler(PASTA_TRABALHO / "analise_sla.log", encoding="utf-8")
arquivo_log.setFormatter(formatter)

# Handler de console
console = logging.StreamHandler()
console.setFormatter(formatter)

if not logger.handlers:
    logger.addHandler(arquivo_log)
    logger.addHandler(console)

# Constantes de data
DATA_ATUAL = datetime.now().strftime("%Y%m%d")
DATA_ONTEM = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
DATA_ATUAL_F = datetime.now().strftime("%d/%m/%Y")
DATA_ONTEM_F = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

VALOR_NAO_ATRIBUIDO = "**Não Atribuido**"


# ==============================================================================
# FUNÇÕES DE MANIPULAÇÃO DE DATAFRAMES
# ==============================================================================
def preencher_nao_atribuido(df: pd.DataFrame) -> pd.DataFrame:
    """Preenche valores nulos ou vazios nas colunas de atribuição com '**Não Atribuido**'."""
    df = df.copy()
    for col in ("Atribuição a", "Atribuído a", "Atribuido a"):
        if col in df.columns:
            df[col] = df[col].replace(r"^\s*$", pd.NA, regex=True).fillna(VALOR_NAO_ATRIBUIDO)
    return df


def agrupar_por(df: pd.DataFrame, coluna: str) -> pd.DataFrame:
    """Agrupa o DataFrame pela coluna especificada e adiciona a linha 'TOTAL GERAL'."""
    aliases = {
        "CI afetado": ["IC afetado", "CI afetado"],
        "IC afetado": ["IC afetado", "CI afetado"],
        "Atribuição a": ["Atribuição a", "Atribuído a"],
        "Atribuído a": ["Atribuído a", "Atribuição a"],
    }

    coluna_origem = coluna
    if coluna not in df.columns and coluna in aliases:
        for alias in aliases[coluna]:
            if alias in df.columns:
                coluna_origem = alias
                break

    if coluna_origem not in df.columns or df.empty:
        if coluna_origem not in df.columns:
            logger.warning(f"Coluna '{coluna}' não encontrada no DataFrame.")
        return pd.DataFrame({
            coluna_origem: ["TOTAL GERAL"],
            "Total": [0]
        })

    # Para colunas de atribuição, preenche vazios com '**Não Atribuido**'
    if coluna_origem in ("Atribuição a", "Atribuído a", "Atribuido a"):
        serie = df[coluna_origem].replace(r"^\s*$", pd.NA, regex=True).fillna(VALOR_NAO_ATRIBUIDO)
    else:
        serie = df[coluna_origem].dropna()

    if serie.empty:
        return pd.DataFrame({
            coluna_origem: ["TOTAL GERAL"],
            "Total": [0]
        })

    resultado = serie.groupby(serie).size().reset_index(name="Total")
    resultado.columns = [coluna_origem, "Total"]
    total_geral = int(resultado["Total"].sum()) if not resultado.empty else 0

    linha_total = pd.DataFrame({
        coluna_origem: ["TOTAL GERAL"],
        "Total": [total_geral]
    })

    return pd.concat([resultado, linha_total], axis=0, ignore_index=True)


# ==============================================================================
# FUNÇÕES DE ESTILIZAÇÃO E FORMATO DO EXCEL
# ==============================================================================
def aplicar_estilos_excel(caminho_arquivo: Path, aba: str) -> None:
    """Aplica o cabeçalho estilizado (Azul escuro com texto branco) na primeira linha."""
    wb = load_workbook(caminho_arquivo)
    if aba not in wb.sheetnames:
        wb.close()
        return

    ws = wb[aba]

    fill_titulo = PatternFill(fill_type="solid", start_color="1F4E79", end_color="1F4E79")
    font_titulo = Font(color="FFFFFF", bold=True)
    borda_fina = Side(border_style="thin", color="000000")
    borda_titulo = Border(top=borda_fina, bottom=borda_fina, left=borda_fina, right=borda_fina)

    for cell in ws[1]:
        cell.fill = fill_titulo
        cell.font = font_titulo
        cell.border = borda_titulo

    wb.save(caminho_arquivo)
    wb.close()


def colocar_cores_celulas(caminho_arquivo: Path, nome_aba: str, nome_coluna: str) -> Dict[str, int]:
    """Aplica cores nas células baseando-se na meta de SLA e contabiliza a distribuição."""
    wb = load_workbook(caminho_arquivo)
    ws = wb[nome_aba]

    valores_sla = {"otimo": 0, "medio": 0, "atencao": 0, "fora": 0, "TOTAL": 0}

    cores = {
        "verde": PatternFill("solid", start_color="4CA154", end_color="4CA154"),
        "amarelo": PatternFill("solid", start_color="D77841", end_color="D77841"),
        "laranja": PatternFill("solid", start_color="F4C142", end_color="F4C142"),
        "vermelho": PatternFill("solid", start_color="E63223", end_color="E63223"),
    }

    col_idx = next((cell.column for cell in ws[1] if cell.value == nome_coluna), None)

    if col_idx is None:
        wb.close()
        raise ValueError(f"Coluna '{nome_coluna}' não encontrada na aba '{nome_aba}'.")

    for row in range(2, ws.max_row + 1):
        celula = ws.cell(row=row, column=col_idx)
        valor = celula.value

        if valor is None:
            continue

        if valor < 50:
            celula.fill = cores["verde"]
            valores_sla["otimo"] += 1
        elif valor < 75:
            celula.fill = cores["amarelo"]
            valores_sla["medio"] += 1
        elif valor < 100:
            celula.fill = cores["laranja"]
            valores_sla["atencao"] += 1
        else:
            celula.fill = cores["vermelho"]
            valores_sla["fora"] += 1

    valores_sla["TOTAL"]= valores_sla["otimo"] + valores_sla["medio"] + valores_sla["atencao"] + valores_sla["fora"]
    wb.save(caminho_arquivo)
    wb.close()

    logger.info(f"Cores aplicadas com sucesso no arquivo: {caminho_arquivo.name}")
    return valores_sla


def aplicar_cores_sla_totais(caminho_arquivo: Path, aba: str, valores_sla: Dict[str, int]) -> None:
    """Preenche a tabela resumo de SLA na aba 'SLA'."""
    wb = load_workbook(caminho_arquivo)

    if "SLA" in wb.sheetnames:
        ws = wb["SLA"]
    else:
        ws = wb.worksheets[1]
        ws.title = "SLA"

    if aba in {"INT_aberto", "DOM_aberto", "CSATH_aberto"}:
        _aplicar_sla_resumo(ws, valores_sla)
    else:
        logger.warning(f"Aba '{aba}' não é elegível para o resumo SLA.")

    wb.save(caminho_arquivo)
    wb.close()
    logger.info(f"Resumo SLA atualizado com sucesso no arquivo '{caminho_arquivo.name}'")


def _aplicar_sla_resumo(ws, valores_sla: Dict[str, int]) -> None:
    """Função auxiliar para estruturar os dados da tabela de resumo de SLA."""
    borda = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )
    fill_cabecalho = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    font_cabecalho = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    ultima_linha = ws.max_row + 3

    letraA = f"A{ultima_linha}"
    letraB = f"B{ultima_linha}"
    # Cabeçalho
    for col, texto in [(letraA, "SLA"), (letraB, "Total")]:
        ws[col].value = texto
        ws[col].fill = fill_cabecalho
        ws[col].font = font_cabecalho
        ws[col].alignment = align_center


    configuracoes = [
        (ultima_linha + 1, "Dentro do SLA", valores_sla.get("otimo", 0), "4ca154"),
        (ultima_linha + 2, "Em andamento", valores_sla.get("medio", 0), "d77841"),
        (ultima_linha + 3, "Em atenção", valores_sla.get("atencao", 0), "f4c142"),
        (ultima_linha + 4, "Fora do SLA", valores_sla.get("fora", 0), "e63223"),
        (ultima_linha + 5, "TOTAL GERAL", valores_sla.get("TOTAL", 0), "FFFFFF"),
    ]

    for linha, texto, valor, cor in configuracoes:
        fill_linha = PatternFill(start_color=cor, end_color=cor, fill_type="solid")

        ws[f"A{linha}"].value = texto
        ws[f"B{linha}"].value = valor

        for col_letter in ["A", "B"]:
            cell = ws[f"{col_letter}{linha}"]
            cell.fill = fill_linha
            cell.alignment = align_center

    for row in ws.iter_rows(min_row=ultima_linha +1 , max_row=ultima_linha + 5 , min_col=1, max_col=2):
        for cell in row:
            cell.border = borda
            
# ==============================================================================
# GERAÇÃO DE RELATÓRIOS
# ==============================================================================
def gerar_relatorio_encerrados(caminho_arquivo: Path, aba1: str, aba2: str, df_filtrado: pd.DataFrame) -> None:
    relatorio = ExcelReport(arquivo=str(caminho_arquivo), sheet_name=aba2, linhas_em_branco=2)
    coluna_ic = next(
        (coluna for coluna in ("IC afetado", "CI afetado") if coluna in df_filtrado.columns),
        "IC afetado",
    )
    df_agrupa1 = agrupar_por(df_filtrado, coluna_ic)
    df_agrupa2 = agrupar_por(df_filtrado, "Resolvido por")

    relatorio.gerar([df_agrupa1, df_agrupa2])
    aplicar_estilos_excel(caminho_arquivo, aba1)


def gerar_relatorio_aberto(caminho_arquivo: Path, aba1: str, aba2: str, df_filtrado: pd.DataFrame) -> None:
    relatorio = ExcelReport(arquivo=str(caminho_arquivo), sheet_name=aba2, linhas_em_branco=2)
    coluna_ic = next(
        (coluna for coluna in ("IC afetado", "CI afetado") if coluna in df_filtrado.columns),
        "IC afetado",
    )
    df_agrupa1 = agrupar_por(df_filtrado, coluna_ic)
    coluna_atribuicao = next(
        (coluna for coluna in ("Atribuição a", "Atribuído a") if coluna in df_filtrado.columns),
        "Atribuído a",
    )
    df_agrupa2 = agrupar_por(df_filtrado, coluna_atribuicao)

    relatorio.gerar([df_agrupa1, df_agrupa2])
    aplicar_estilos_excel(caminho_arquivo, aba1)
    
    #gera a metricas para o campo CI afetado
    gerar_ci_json(caminho_arquivo, coluna_ci=coluna_ic)


def gerar_relatorio_aberto_sctask(caminho_arquivo: Path, aba1: str, aba2: str, df_filtrado: pd.DataFrame) -> None:
    relatorio = ExcelReport(arquivo=str(caminho_arquivo), sheet_name=aba2, linhas_em_branco=2)
    df_agrupa1 = agrupar_por(df_filtrado, "Grupo de atribuição")
    coluna_atribuicao = next(
        (coluna for coluna in ("Atribuição a", "Atribuído a") if coluna in df_filtrado.columns),
        "Atribuído a",
    )
    df_agrupa2 = agrupar_por(df_filtrado, coluna_atribuicao)

    relatorio.gerar([df_agrupa1, df_agrupa2])
    aplicar_estilos_excel(caminho_arquivo, aba1)


# ==============================================================================
# PROCESSAMENTO DE PLANILHAS
# ==============================================================================
def cruzar_planilhas(arquivo1: Path, arquivo2: Path, arquivo_saida: Path, filtro: str, aba1: str, aba2: str) -> None:
    df_plan1 = pd.read_excel(arquivo1)
    df_plan2 = pd.read_excel(arquivo2)

    df_plan1.rename(columns={df_plan1.columns[0]: 'Numero', df_plan1.columns[2]: 'Aberto por'}, inplace=True)
    df_plan2.rename(columns={df_plan2.columns[0]: 'Numero', df_plan2.columns[6]: '% SLA Resolution'}, inplace=True)

    df_plan2 = df_plan2.iloc[:, [6, 0]]
    df_filtrado = df_plan1[df_plan1.iloc[:, 5].astype(str).str.contains(filtro, na=False, regex=True)]

    df_plan3 = pd.merge(df_plan2, df_filtrado, on='Numero', how='right')
    df_plan3.insert(loc=0, column='% SLA Response', value='')
    df_plan3['% SLA Resolution'] = df_plan3['% SLA Resolution'].fillna(0)
    df_plan3 = df_plan3.query("Estado != 'Fechados' and Estado != 'Cancelado'")
    df_plan3 = preencher_nao_atribuido(df_plan3)

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        df_plan3.to_excel(writer, sheet_name=aba1, index=False)
        pd.DataFrame().to_excel(writer, sheet_name=aba2, index=False)

    aplicar_estilos_excel(arquivo_saida, aba1)
    logger.info(f"Arquivo '{arquivo_saida.name}' cruzado com sucesso!")

    gerar_relatorio_aberto(arquivo_saida, aba1, aba2, df_plan3)


def planilhas_encerrados(arquivo1: Path, arquivo_saida: Path, filtro: str, aba1: str, aba2: str) -> None:
    df_plan1 = pd.read_excel(arquivo1)
    df_plan1.rename(columns={df_plan1.columns[2]: 'Aberto por'}, inplace=True)

    df_filtrado = df_plan1[df_plan1.iloc[:, 5] == filtro]
    df_filtrado = preencher_nao_atribuido(df_filtrado)

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, sheet_name=aba1, index=False)
        pd.DataFrame().to_excel(writer, sheet_name=aba2, index=False)

    aplicar_estilos_excel(arquivo_saida, aba1)
    logger.info(f"Arquivo '{arquivo_saida.name}' gerado com sucesso!")

    gerar_relatorio_encerrados(arquivo_saida, aba1, aba2, df_filtrado)


def planilhas_encerrados_aberto_sctask(arquivo1: Path, arquivo_saida: Path, aba1: str, aba2: str) -> None:
    df_plan1 = pd.read_excel(arquivo1)
    df_plan1 = preencher_nao_atribuido(df_plan1)

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        df_plan1.to_excel(writer, sheet_name=aba1, index=False)
        pd.DataFrame().to_excel(writer, sheet_name=aba2, index=False)

    gerar_relatorio_aberto_sctask(arquivo_saida, aba1, aba2, df_plan1)
    aplicar_estilos_excel(arquivo_saida, aba1)
    logger.info(f"Arquivo SCTASK '{arquivo_saida.name}' criado com sucesso!")


# ==============================================================================
# FLUXO PRINCIPAL
# ==============================================================================
def processar_configuracoes() -> None:
    caminho_config = PASTA_TRABALHO / "configuracao.json"

    if not caminho_config.exists():
        logger.error(f"Arquivo de configuração não encontrado: {caminho_config}")
        return

    with open(caminho_config, "r", encoding="utf-8") as f:
        configuracoes = json.load(f)

    for item in configuracoes:
        arquivo1 = PASTA_TRABALHO / item["ArquivoEntrada1"]
        arquivo2 = PASTA_TRABALHO / item["ArquivoEntrada2"]
        arquivo_saida = PASTA_TRABALHO / f"{DATA_ATUAL}_{item['arquivoSaida']}"

        filtro = item["Filtro"]
        junta_arq = item.get("JuntaArq", False)
        aba1 = item["NovaAba1"]
        aba2 = item["NovaAba2"]

        # Estrutura padrão de SLA caso a contagem de células não seja executada
        valores_sla_padrao = {"otimo": 0, "medio": 0, "atencao": 0, "fora": 0, "TOTAL":0}

        if "SCTASK" in arquivo1.name:
            planilhas_encerrados_aberto_sctask(arquivo1, arquivo_saida, aba1, aba2)
            aplicar_cores_sla_totais(arquivo_saida, aba1, valores_sla_padrao)

        elif junta_arq:
            planilhas_encerrados(arquivo1, arquivo_saida, filtro, aba1, aba2)
            # Caso queira aplicar cálculo de cores no modo junta_arq, descomente abaixo:
            # valores_sla = colocar_cores_celulas(arquivo_saida, aba1, '% SLA Resolution')
            aplicar_cores_sla_totais(arquivo_saida, aba1, valores_sla_padrao)

        else:
            cruzar_planilhas(arquivo1, arquivo2, arquivo_saida, filtro, aba1, aba2)
            valores_sla = colocar_cores_celulas(arquivo_saida, aba1, '% SLA Resolution')
            aplicar_cores_sla_totais(arquivo_saida, aba1, valores_sla)

def gerar_ci_json(
    arquivo_excel,
    arquivo_json=None,
    coluna_ci=None,
):
    """
    Lê a coluna CI do Excel, remove duplicidades
    e gera um arquivo JSON.

    Args:
        arquivo_excel (str | Path): Caminho do arquivo Excel.
        arquivo_json (str | Path, optional): Caminho do arquivo JSON de saída.
        coluna_ci (str, optional): Nome da coluna contendo os CIs.
    """
    if arquivo_json is None:
        arquivo_json = PASTA_TRABALHO / "ci.json"

    # Lê todas as abas
    abas = pd.read_excel(
        arquivo_excel,
        sheet_name=[0] #lendo so a aba selecionada, para todoas substituir por None
    )

    cis = set()

    # Carrega CIs já existentes no arquivo JSON para acumular entre relatórios
    if Path(arquivo_json).exists():
        try:
            with open(arquivo_json, "r", encoding="utf-8") as f:
                dados_existentes = json.load(f)
                if isinstance(dados_existentes, list):
                    for item in dados_existentes:
                        if isinstance(item, dict) and "CI" in item:
                            cis.add(item["CI"])
        except Exception:
            pass

    # Coleta CIs únicos de todas as abas
    for _, df in abas.items():
        col = coluna_ci
        if col is None or col not in df.columns:
            col = next(
                (c for c in ("IC afetado", "CI afetado") if c in df.columns),
                None,
            )
        if col and col in df.columns:
            cis.update(
                df[col]
                .dropna()
                .astype(str)
                .unique()
            )

    dados = [
        {
            "CI": ci,
            "otimo": 0,
            "medio": 0,
            "atencao": 0,
            "fora": 0,
            "total": 0
        }
        for ci in sorted(cis)
    ]

    with open(arquivo_json, "w", encoding="utf-8") as f:
        json.dump(
            dados,
            f,
            ensure_ascii=False,
            indent=4
        )

    logger.info(
        f"Arquivo '{Path(arquivo_json).name}' atualizado com {len(dados)} CIs."
    )


def criar_planilha_total() -> None:
    arquivo_saida = gts.gerar_planilha_total(pasta=PASTA_TRABALHO, prefixo=DATA_ATUAL)
    logger.info(f"Planilha resumo final '{arquivo_saida.name}' criada com sucesso!")


def main() -> None:
    logger.info("Iniciando processamento dos relatórios de SLA...")
    processar_configuracoes()
    criar_planilha_total()
    logger.info("Processamento finalizado com sucesso!")


if __name__ == "__main__":
    main()