import pandas as pd
import json
import re
import sys
import subprocess
from datetime import datetime
from pathlib import Path

# ==========================================================
# BASE DIR (funciona no .py e no .exe)
# ==========================================================
def descobrir_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent.parent

    for _ in range(6):
        if (base / "excel").exists() and (base / "site").exists():
            return base
        base = base.parent

    return Path(__file__).resolve().parent.parent


BASE_DIR = descobrir_base_dir()

EXCEL_2026 = BASE_DIR / "excel" / "PEDIDOS_2026.xlsx"
EXCEL_2025 = BASE_DIR / "excel" / "PEDIDOS_2025.xlsx"

DADOS_DIR_1 = BASE_DIR / "dados"
DADOS_DIR_2 = BASE_DIR / "site" / "dados"

# ==========================================================
# UTIL
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
# CONVERSÃO NÚMEROS EXCEL
# ==========================================================
EXCEL_EPOCH = datetime(1899, 12, 30)

def excel_datetime_para_numero(dt: datetime) -> float:
    delta = dt - EXCEL_EPOCH
    return delta.days + (delta.seconds / 86400)


def limpar_numero(v) -> float:
    if pd.isna(v):
        return 0.0

    if isinstance(v, (int, float)):
        return float(v)

    if isinstance(v, (pd.Timestamp, datetime)):
        dt = v.to_pydatetime() if isinstance(v, pd.Timestamp) else v
        if dt.year in (1899, 1900, 1901):
            return float(excel_datetime_para_numero(dt))
        return 0.0

    s = str(v).strip()
    s = re.sub(r"[^0-9,.\-]", "", s)

    if s in ("", "-", ",", ".", ",-", ".-"):
        return 0.0

    if "." in s and "," in s:
        return float(s.replace(".", "").replace(",", "."))

    if "," in s and "." not in s:
        return float(s.replace(",", "."))

    if "." in s and "," not in s:
        partes = s.split(".")
        if len(partes[-1]) == 3 and len(partes) > 1:
            s = s.replace(".", "")
        return float(s)

    return float(s)


# ==========================================================
# CARREGAR PLANILHA 2025/2026
# ==========================================================
def carregar_planilha(caminho: Path) -> pd.DataFrame:
    df = pd.read_excel(caminho)
    df = normalizar_colunas(df)

    col_data = achar_coluna(df, ["DATA", "DT", "DATA PEDIDO"])
    col_tipo = achar_coluna(df, ["TIPO DE PEDIDO", "TIPO"])
    col_pedido = achar_coluna(df, ["PEDIDO"])
    col_valor_ipi = achar_coluna(df, ["VALOR COM IPI"])
    col_kg = achar_coluna(df, ["KG"])
    col_m2 = achar_coluna(df, ["TOTAL M2", "M2"])

    df[col_data] = pd.to_datetime(df[col_data], errors="coerce", dayfirst=True)
    df = df[df[col_data].notna()]

    df[col_tipo] = df[col_tipo].astype(str).str.upper()
    df = df[df[col_tipo] == "NORMAL"]

    df[col_valor_ipi] = df[col_valor_ipi].apply(limpar_numero)
    df[col_kg] = df[col_kg].apply(limpar_numero)
    df[col_m2] = df[col_m2].apply(limpar_numero)

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
# DEFINIR PERÍODO — usa última data real da planilha 2026
# ==========================================================
def definir_periodo(df_2026, data_inicio=None, data_fim=None):
    if data_inicio and data_fim:
        return (
            datetime.strptime(data_inicio, "%d/%m/%Y"),
            datetime.strptime(data_fim, "%d/%m/%Y")
        )

    ultima = df_2026["DATA"].max()
    return (
        ultima.replace(day=1),
        ultima
    )


# ==========================================================
# GERAR RESUMO
# ==========================================================
def resumo(df, inicio, fim):
    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)]

    pedidos = d["PEDIDO"].nunique()
    fat = float(d["VALOR COM IPI"].sum())
    kg = float(d["KG"].sum())
    m2 = float(d["TOTAL M2"].sum())

    return {
        "pedidos": pedidos,
        "fat": fat,
        "kg": kg,
        "m2": m2,
        "ticket": fat / pedidos if pedidos else 0,
        "preco_kg": fat / kg if kg else 0,
        "preco_m2": fat / m2 if m2 else 0,
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y"),
    }


# ==========================================================
# SALVAR JSON
# ==========================================================
def salvar_json(nome, payload):
    DADOS_DIR_2.mkdir(parents=True, exist_ok=True)
    with open(DADOS_DIR_2 / nome, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)


# ==========================================================
# 🔥 PUSH AUTOMÁTICO PARA GITHUB
# ==========================================================
def push_github():
    try:
        subprocess.run(["git", "add", "."], cwd=BASE_DIR, check=True)
        subprocess.run(["git", "commit", "-m", "Atualização automática painel"], cwd=BASE_DIR, check=True)
        subprocess.run(["git", "push", "origin", "main"], cwd=BASE_DIR, check=True)
        return True
    except Exception as e:
        return False


# ==========================================================
# MAIN — usado pelo .EXE
# ==========================================================
def main(data_inicio=None, data_fim=None):
    df_2026 = carregar_planilha(EXCEL_2026)
    df_2025 = carregar_planilha(EXCEL_2025)

    inicio, fim = definir_periodo(df_2026, data_inicio, data_fim)
    inicio_ant = inicio.replace(year=inicio.year - 1)
    fim_ant = fim.replace(year=fim.year - 1)

    atual = resumo(df_2026, inicio, fim)
    anterior = resumo(df_2025, inicio_ant, fim_ant)

    # salvar JSONs normais aqui...
    salvar_json("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"]/anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": atual["inicio"],
        "data_atual": atual["fim"],
        "inicio_mes_anterior": anterior["inicio"],
        "data_ano_anterior": anterior["fim"]
    })

    salvar_json("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"]/anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    })

    salvar_json("kpi_kg_total.json", {
        "atual": atual["kg"],
        "ano_anterior": anterior["kg"],
        "variacao": ((atual["kg"]/anterior["kg"]) - 1) * 100 if anterior["kg"] else 0
    })

    salvar_json("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": ((atual["ticket"]/anterior["ticket"]) - 1) * 100 if anterior["ticket"] else 0
    })

    salvar_json("kpi_preco_medio.json", {
        "atual": {
            "preco_medio_kg": atual["preco_kg"],
            "preco_medio_m2": atual["preco_m2"],
            "total_kg": atual["kg"],
            "total_m2": atual["m2"],
            "data": atual["fim"],
        },
        "ano_anterior": {
            "preco_medio_kg": anterior["preco_kg"],
            "preco_medio_m2": anterior["preco_m2"],
            "total_kg": anterior["kg"],
            "total_m2": anterior["m2"],
            "data": anterior["fim"],
        }
    })

    # PUSH AUTOMÁTICO
    push_ok = push_github()

    return {
        "atual": atual,
        "anterior": anterior,
        "push": push_ok
    }


if __name__ == "__main__":
    main()
