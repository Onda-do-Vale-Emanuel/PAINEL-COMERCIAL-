import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ---------------------------------------------------------
# Função definitiva para limpar números (R$ 1.234,56)
# ---------------------------------------------------------
def limpar_numero(valor):
    if pd.isna(valor):
        return 0.0
    v = str(valor).strip()
    v = re.sub(r"[^0-9,.-]", "", v)

    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0

    if "." in v and "," in v:
        return float(v.replace(".", "").replace(",", "."))

    if "," in v:
        return float(v.replace(",", "."))

    return float(v)


# ---------------------------------------------------------
# Carregar Excel
# ---------------------------------------------------------
def carregar_excel():
    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.strip().str.upper()
    obrig = ["DATA", "VALOR COM IPI", "KG", "TOTAL M2"]

    for c in obrig:
        if c not in df.columns:
            raise Exception(f"❌ Coluna ausente: {c}")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    df["VALOR COM IPI"] = df["VALOR COM IPI"].apply(limpar_numero)
    df["KG"] = df["KG"].apply(limpar_numero)
    df["TOTAL M2"] = df["TOTAL M2"].apply(limpar_numero)

    return df


# ---------------------------------------------------------
# Definir período correto
# 01/MM/AAAA até última data do mês
# ---------------------------------------------------------
def obter_periodo(df):
    ultima = df["DATA"].max()
    primeira = ultima.replace(day=1)

    return primeira, ultima


# ---------------------------------------------------------
# Calcular tudo do painel
# ---------------------------------------------------------
def calcular(df):
    primeira, ultima = obter_periodo(df)
    ano_atual = ultima.year
    ano_anterior = ano_atual - 1

    # Período atual
    df_atual = df[(df["DATA"] >= primeira) & (df["DATA"] <= ultima)]

    # Período ano anterior
    primeira_ant = primeira.replace(year=ano_anterior)
    ultima_ant = ultima.replace(year=ano_anterior)

    df_ant = df[(df["DATA"] >= primeira_ant) & (df["DATA"] <= ultima_ant)]

    # ---- SOMAS ----
    total_valor = df_atual["VALOR COM IPI"].sum()
    total_kg = df_atual["KG"].sum()
    total_m2 = df_atual["TOTAL M2"].sum()
    qtd = len(df_atual)

    total_valor_ant = df_ant["VALOR COM IPI"].sum()
    total_kg_ant = df_ant["KG"].sum()
    total_m2_ant = df_ant["TOTAL M2"].sum()
    qtd_ant = len(df_ant)

    # ---- TICKET ----
    ticket_atual = total_valor / qtd if qtd else 0
    ticket_ant = total_valor_ant / qtd_ant if qtd_ant else 0

    # ---- PREÇO MÉDIO ----
    preco_kg = round(total_valor / total_kg, 2) if total_kg else 0
    preco_m2 = round(total_valor / total_m2, 2) if total_m2 else 0

    return {
        "fat": {
            "atual": total_valor,
            "ano_anterior": total_valor_ant,
            "variacao": ((total_valor / total_valor_ant) - 1) * 100 if total_valor_ant else 0,
            "data_atual": ultima.strftime("%d/%m/%Y"),
            "data_ano_anterior": ultima_ant.strftime("%d/%m/%Y"),
        },
        "qtd": {
            "atual": qtd,
            "ano_anterior": qtd_ant,
            "variacao": ((qtd / qtd_ant) - 1) * 100 if qtd_ant else 0,
        },
        "kg": {
            "atual": total_kg,
            "ano_anterior": total_kg_ant,
            "variacao": ((total_kg / total_kg_ant) - 1) * 100 if total_kg_ant else 0,
        },
        "ticket": {
            "atual": ticket_atual,
            "ano_anterior": ticket_ant,
            "variacao": ((ticket_atual / ticket_ant) - 1) * 100 if ticket_ant else 0,
        },
        "preco": {
            "preco_medio_kg": preco_kg,
            "preco_medio_m2": preco_m2,
            "total_kg": total_kg,
            "total_m2": total_m2,
            "data": ultima.strftime("%d/%m/%Y"),
        },
        "primeira": primeira.strftime("%d/%m/%Y"),
        "ultima": ultima.strftime("%d/%m/%Y"),
        "primeira_ant": primeira_ant.strftime("%d/%m/%Y"),
        "ultima_ant": ultima_ant.strftime("%d/%m/%Y"),
    }


# ---------------------------------------------------------
# Salvar JSONs
# ---------------------------------------------------------
def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------
# Executar
# ---------------------------------------------------------
if __name__ == "__main__":
    df = carregar_excel()
    res = calcular(df)

    salvar("kpi_faturamento.json", res["fat"])
    salvar("kpi_quantidade_pedidos.json", res["qtd"])
    salvar("kpi_kg_total.json", res["kg"])
    salvar("kpi_ticket_medio.json", res["ticket"])
    salvar("kpi_preco_medio.json", res["preco"])

    print("✓ JSON gerados!")
    print("📌 Exibindo no site:")
    print("Ano atual:", res["primeira"], "→", res["ultima"])
    print("Ano anterior:", res["primeira_ant"], "→", res["ultima_ant"])
