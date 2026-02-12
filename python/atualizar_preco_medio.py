import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ======================================================
# LIMPAR NÚMEROS
# ======================================================
def limpar_numero(v):

    if pd.isna(v):
        return 0.0

    if isinstance(v, datetime):
        base = datetime(1899, 12, 30)
        dias = (v - base).days
        fracao = (v - base).seconds / 86400
        return round(dias + fracao, 2)

    v = str(v).strip()
    v = re.sub(r"[^0-9,.-]", "", v)

    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0

    if "." in v and "," in v:
        return float(v.replace(".", "").replace(",", "."))

    if "," in v:
        return float(v.replace(",", "."))

    return float(v)


# ======================================================
# CARREGAR
# ======================================================
def carregar():

    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.upper().str.strip()

    for col in ["VALOR COM IPI", "KG", "TOTAL M2"]:
        if col in df.columns:
            df[col] = df[col].apply(limpar_numero)

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    df["PEDIDO_NUM"] = df["PEDIDO"].apply(limpar_numero)
    df = df[(df["PEDIDO_NUM"] >= 30000) & (df["PEDIDO_NUM"] <= 50000)]

    return df


# ======================================================
# RESUMO
# ======================================================
def resumo(df, inicio, fim):

    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)]

    total_valor = d["VALOR COM IPI"].sum()
    total_kg = d["KG"].sum()
    total_m2 = d["TOTAL M2"].sum()

    preco_kg = total_valor / total_kg if total_kg else 0
    preco_m2 = total_valor / total_m2 if total_m2 else 0

    return {
        "preco_medio_kg": round(preco_kg, 2),
        "preco_medio_m2": round(preco_m2, 2)
    }


# ======================================================
# EXECUÇÃO
# ======================================================
if __name__ == "__main__":

    print("=====================================")
    print("Atualizando PREÇO MÉDIO")
    print("=====================================")

    df = carregar()

    hoje = datetime.today()

    inicio_atual = hoje.replace(day=1, hour=0, minute=0, second=0)
    fim_atual = hoje

    inicio_ant = inicio_atual.replace(year=hoje.year - 1)
    fim_ant = fim_atual.replace(year=hoje.year - 1)

    atual = resumo(df, inicio_atual, fim_atual)
    anterior = resumo(df, inicio_ant, fim_ant)

    dados = {
        "atual": {
            "preco_medio_kg": atual["preco_medio_kg"],
            "preco_medio_m2": atual["preco_medio_m2"],
            "data": fim_atual.strftime("%d/%m/%Y")
        },
        "ano_anterior": {
            "preco_medio_kg": anterior["preco_medio_kg"],
            "preco_medio_m2": anterior["preco_medio_m2"],
            "data": fim_ant.strftime("%d/%m/%Y")
        }
    }

    with open("dados/kpi_preco_medio.json", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    with open("site/dados/kpi_preco_medio.json", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    print("✓ PREÇO MÉDIO ATUALIZADO ATÉ HOJE")
    print("=====================================")
