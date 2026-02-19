import pandas as pd
import json
import re
import sys
from datetime import datetime
from pathlib import Path

# ==========================================================
# BASE DIR (funciona no .py e no .exe)
# ==========================================================
def descobrir_base_dir() -> Path:
    # Se estiver empacotado (exe), sys.executable aponta para o .exe
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent.parent

    # Garante que estamos na raiz do projeto (onde existe /excel e /site)
    # Se não achar, sobe alguns níveis.
    for _ in range(6):
        if (base / "excel").exists() and (base / "site").exists():
            return base
        base = base.parent

    # fallback
    return Path(__file__).resolve().parent.parent


BASE_DIR = descobrir_base_dir()

EXCEL_2026 = BASE_DIR / "excel" / "PEDIDOS_2026.xlsx"
EXCEL_2025 = BASE_DIR / "excel" / "PEDIDOS_2025.xlsx"

DADOS_DIR_1 = BASE_DIR / "dados"          # (se existir no projeto)
DADOS_DIR_2 = BASE_DIR / "site" / "dados" # (usado pelo site)

# ==========================================================
# UTIL: normalizar nomes de colunas
# ==========================================================
def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def achar_coluna(df: pd.DataFrame, candidatos):
    cols = set(df.columns)
    for c in candidatos:
        if c in cols:
            return c
    return None

# ==========================================================
# 🔥 CONVERSÃO "NÚMERO BRASILEIRO" + CORREÇÃO DO "1900"
# Ex: KG veio como data 29/10/1900 01:12:00 -> vira 303,05
# ==========================================================
EXCEL_EPOCH = datetime(1899, 12, 30)

def excel_datetime_para_numero(dt: datetime) -> float:
    # Converte datetime do Excel para serial (dias) com fração do dia
    delta = dt - EXCEL_EPOCH
    return delta.days + (delta.seconds / 86400)

def limpar_numero(v) -> float:
    if pd.isna(v):
        return 0.0

    # Se já for número
    if isinstance(v, (int, float)):
        try:
            return float(v)
        except:
            return 0.0

    # Timestamp / datetime (o caso do "1900-10-29 01:12:00")
    if isinstance(v, (pd.Timestamp, datetime)):
        dt = v.to_pydatetime() if isinstance(v, pd.Timestamp) else v
        # Se for aquele intervalo típico do Excel 1900 (erro de formatação):
        # converte para serial (ex: 303.05)
        if dt.year in (1899, 1900, 1901):
            return float(excel_datetime_para_numero(dt))
        # Se não for 1900, não faz sentido ser KG/valor -> retorna 0
        return 0.0

    s = str(v).strip()

    # Alguns casos podem vir assim: "1900-10-29011200" (mistura)
    # tenta extrair só números e separadores
    s = re.sub(r"[^0-9,.\-]", "", s)

    if s in ("", "-", ",", ".", ",-", ".-"):
        return 0.0

    # formato BR: 1.234,56
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0

    # formato com vírgula: 123,45
    if "," in s and "." not in s:
        try:
            return float(s.replace(",", "."))
        except:
            return 0.0

    # formato com ponto: 1234.56 ou 1.234 (milhar)
    if "." in s and "," not in s:
        partes = s.split(".")
        # se última parte tiver 3 dígitos, provavelmente milhar
        if len(partes[-1]) == 3 and len(partes) > 1:
            s = s.replace(".", "")
        try:
            return float(s)
        except:
            return 0.0

    try:
        return float(s)
    except:
        return 0.0

# ==========================================================
# CARREGAR E LIMPAR PLANILHA (2025/2026)
# ==========================================================
def carregar_planilha(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    df = pd.read_excel(caminho, sheet_name=0)
    df = normalizar_colunas(df)

    # achar colunas (tolerante a variações)
    col_data = achar_coluna(df, ["DATA", "DT", "DATA PEDIDO", "EMISSÃO", "EMISSAO"])
    if not col_data:
        raise KeyError(f"Não encontrei a coluna de DATA em {caminho.name}. Colunas: {list(df.columns)}")

    col_tipo = achar_coluna(df, ["TIPO DE PEDIDO", "TIPO_PEDIDO", "TIPO"])
    if not col_tipo:
        raise KeyError(f"Não encontrei a coluna TIPO DE PEDIDO em {caminho.name}. Colunas: {list(df.columns)}")

    col_pedido = achar_coluna(df, ["PEDIDO", "Nº PEDIDO", "NUM PEDIDO", "NÚMERO PEDIDO", "NUMERO PEDIDO"])
    if not col_pedido:
        raise KeyError(f"Não encontrei a coluna PEDIDO em {caminho.name}. Colunas: {list(df.columns)}")

    col_valor_ipi = achar_coluna(df, ["VALOR COM IPI", "VLR COM IPI", "VALOR C/ IPI"])
    if not col_valor_ipi:
        raise KeyError(f"Não encontrei a coluna VALOR COM IPI em {caminho.name}. Colunas: {list(df.columns)}")

    col_kg = achar_coluna(df, ["KG", "PESO", "PESO KG", "TOTAL KG"])
    if not col_kg:
        raise KeyError(f"Não encontrei a coluna KG em {caminho.name}. Colunas: {list(df.columns)}")

    col_m2 = achar_coluna(df, ["TOTAL M2", "TOTAL M²", "M2", "M²", "TOTAL_M2"])
    if not col_m2:
        raise KeyError(f"Não encontrei a coluna TOTAL M2 em {caminho.name}. Colunas: {list(df.columns)}")

    # datas
    df[col_data] = pd.to_datetime(df[col_data], errors="coerce", dayfirst=True)
    df = df[df[col_data].notna()]

    # filtro tipo (Normal) robusto
    df[col_tipo] = df[col_tipo].astype(str).str.strip().str.upper()
    df = df[df[col_tipo] == "NORMAL"]

    # limpar numéricos
    df[col_valor_ipi] = df[col_valor_ipi].apply(limpar_numero)
    df[col_kg] = df[col_kg].apply(limpar_numero)
    df[col_m2] = df[col_m2].apply(limpar_numero)

    # guardar colunas padronizadas para o resto do código
    df = df.rename(columns={
        col_data: "DATA",
        col_tipo: "TIPO DE PEDIDO",
        col_pedido: "PEDIDO",
        col_valor_ipi: "VALOR COM IPI",
        col_kg: "KG",
        col_m2: "TOTAL M2",
    })

    return df

# ==========================================================
# PERÍODO: se usuário informar, usa; se não, usa mês da ÚLTIMA DATA REAL do 2026
# (isso evita ficar preso ou pegar "hoje" sem existir na planilha)
# ==========================================================
def definir_periodo(df_2026: pd.DataFrame, data_inicio=None, data_fim=None):
    if data_inicio and data_fim:
        inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
        fim = datetime.strptime(data_fim, "%d/%m/%Y")
    else:
        ultima_data_2026 = df_2026["DATA"].max()
        inicio = ultima_data_2026.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        fim = ultima_data_2026.replace(hour=0, minute=0, second=0, microsecond=0)
    return inicio, fim

# ==========================================================
# RESUMO KPIs
# ==========================================================
def resumo(df: pd.DataFrame, inicio: datetime, fim: datetime):
    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)].copy()

    pedidos = d["PEDIDO"].nunique()  # mais seguro que len(d)
    fat = float(d["VALOR COM IPI"].sum())
    kg = float(d["KG"].sum())
    m2 = float(d["TOTAL M2"].sum())

    ticket = fat / pedidos if pedidos else 0.0
    preco_kg = fat / kg if kg else 0.0
    preco_m2 = fat / m2 if m2 else 0.0

    return {
        "pedidos": int(pedidos),
        "fat": round(fat, 2),
        "kg": round(kg, 2),
        "m2": round(m2, 2),
        "ticket": round(ticket, 2),
        "preco_kg": round(preco_kg, 2),
        "preco_m2": round(preco_m2, 2),
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y"),
    }

# ==========================================================
# SALVAR JSONS (em /site/dados e também em /dados se existir)
# ==========================================================
def salvar_json(nome: str, payload: dict):
    DADOS_DIR_2.mkdir(parents=True, exist_ok=True)
    with open(DADOS_DIR_2 / nome, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    # opcional: manter espelho em /dados se existir
    if DADOS_DIR_1.exists():
        with open(DADOS_DIR_1 / nome, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

# ==========================================================
# MAIN (para usar no .exe e no bat)
# ==========================================================
def main(data_inicio=None, data_fim=None):
    df_2026 = carregar_planilha(EXCEL_2026)
    df_2025 = carregar_planilha(EXCEL_2025)

    inicio, fim = definir_periodo(df_2026, data_inicio, data_fim)

    # ano anterior = mesmo intervalo de datas, trocando o ano
    inicio_ant = inicio.replace(year=inicio.year - 1)
    fim_ant = fim.replace(year=fim.year - 1)

    atual = resumo(df_2026, inicio, fim)
    anterior = resumo(df_2025, inicio_ant, fim_ant)

    # KPI faturamento
    variacao_fat = ((atual["fat"] / anterior["fat"]) - 1) * 100 if anterior["fat"] else 0
    salvar_json("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": variacao_fat,
        "inicio_mes": atual["inicio"],
        "data_atual": atual["fim"],
        "inicio_mes_anterior": anterior["inicio"],
        "data_ano_anterior": anterior["fim"],
    })

    # KPI quantidade pedidos
    variacao_qtd = ((atual["pedidos"] / anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    salvar_json("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": variacao_qtd
    })

    # KPI kg total (site quer sem casas -> você formata no JS; aqui mantemos 2 e no JS exibimos sem)
    variacao_kg = ((atual["kg"] / anterior["kg"]) - 1) * 100 if anterior["kg"] else 0
    salvar_json("kpi_kg_total.json", {
        "atual": atual["kg"],
        "ano_anterior": anterior["kg"],
        "variacao": variacao_kg
    })

    # KPI ticket médio
    variacao_ticket = ((atual["ticket"] / anterior["ticket"]) - 1) * 100 if anterior["ticket"] else 0
    salvar_json("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": variacao_ticket
    })

    # KPI preço médio (FORMATO CERTO PARA O kpis.js: tem atual e ano_anterior)
    salvar_json("kpi_preco_medio.json", {
        "atual": {
            "preco_medio_kg": atual["preco_kg"],
            "preco_medio_m2": atual["preco_m2"],
            "total_kg": round(atual["kg"], 2),
            "total_m2": round(atual["m2"], 2),
            "data": atual["fim"]
        },
        "ano_anterior": {
            "preco_medio_kg": anterior["preco_kg"],
            "preco_medio_m2": anterior["preco_m2"],
            "total_kg": round(anterior["kg"], 2),
            "total_m2": round(anterior["m2"], 2),
            "data": anterior["fim"]
        }
    })

    print("\n=====================================")
    print("ATUALIZAÇÃO CONCLUÍDA COM SUCESSO")
    print("=====================================")
    print(f"Pedidos 2026: {atual['pedidos']}")
    print(f"Pedidos 2025: {anterior['pedidos']}")
    print(f"Faturamento 2026: {atual['fat']}")
    print(f"Faturamento 2025: {anterior['fat']}")
    print(f"Período 2026: {atual['inicio']} até {atual['fim']}")
    print(f"Período 2025: {anterior['inicio']} até {anterior['fim']}")
    print("=====================================\n")


if __name__ == "__main__":
    # modo automático
    main()
