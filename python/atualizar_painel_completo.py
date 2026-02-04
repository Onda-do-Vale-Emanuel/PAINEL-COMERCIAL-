import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ======================================================
# LIMPA NÚMEROS BRASILEIROS
# ======================================================
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

# ======================================================
# LIMPA PEDIDO (🔥 CHAVE DA CORREÇÃO 🔥)
# ======================================================
def limpar_pedido(valor):
    if pd.isna(valor):
        return None
    v = str(valor).strip()
    v = re.sub(r"[^0-9]", "", v)  # remove . , espaços, texto
    return v if v != "" else None

# ======================================================
# CARREGA EXCEL
# ======================================================
def carregar_excel():
    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.strip().str.upper()

    obrig = ["DATA", "PEDIDO", "VALOR COM IPI", "KG", "TOTAL M2"]
    for c in obrig:
        if c not in df.columns:
            raise Exception(f"❌ Coluna obrigatória não encontrada: {c}")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    # 🔥 NORMALIZA PEDIDOS
    df["PEDIDO_LIMPO"] = df["PEDIDO"].apply(limpar_pedido)
    df = df[df["PEDIDO_LIMPO"].notna()]

    df["VALOR COM IPI"] = df["VALOR COM IPI"].apply(limpar_numero)
    df["KG"] = df["KG"].apply(limpar_numero)
    df["TOTAL M2"] = df["TOTAL M2"].apply(limpar_numero)

    return df

# ======================================================
# PERÍODO DO MÊS ATUAL
# ======================================================
def obter_periodo_mes(df):
    ultima = df["DATA"].max()
    mes = ultima.month
    ano = ultima.year

    df_mes = df[(df["DATA"].dt.month == mes) & (df["DATA"].dt.year == ano)]
    primeira = df_mes["DATA"].min()

    return primeira, ultima

# ======================================================
# CÁLCULOS
# ======================================================
def calcular(df):
    primeira, ultima = obter_periodo_mes(df)

    df_atual = df[(df["DATA"] >= primeira) & (df["DATA"] <= ultima)]

    # ✅ AGORA SIM — pedidos reais
    qtd_atual = df_atual["PEDIDO_LIMPO"].nunique()

    total_valor = df_atual["VALOR COM IPI"].sum()
    total_kg = df_atual["KG"].sum()
    total_m2 = df_atual["TOTAL M2"].sum()

    preco_kg = round(total_valor / total_kg, 2) if total_kg else 0
    preco_m2 = round(total_valor / total_m2, 2) if total_m2 else 0

    # ================= ANO ANTERIOR =================
    ano_ant = primeira.year - 1
    primeira_ant = primeira.replace(year=ano_ant)
    ultima_ant = ultima.replace(year=ano_ant)

    df_ant = df[
        (df["DATA"] >= primeira_ant) &
        (df["DATA"] <= ultima_ant)
    ]

    qtd_ant = df_ant["PEDIDO_LIMPO"].nunique()
    total_valor_ant = df_ant["VALOR COM IPI"].sum()
    total_kg_ant = df_ant["KG"].sum()

    ticket_atual = total_valor / qtd_atual if qtd_atual else 0
    ticket_ant = total_valor_ant / qtd_ant if qtd_ant else 0

    return {
        "faturamento": {
            "atual": round(total_valor, 2),
            "ano_anterior": round(total_valor_ant, 2),
            "variacao": ((total_valor / total_valor_ant) - 1) * 100 if total_valor_ant else 0,
            "data": ultima.strftime("%d/%m/%Y"),
            "inicio_mes": primeira.strftime("%d/%m/%Y"),
            "data_ano_anterior": ultima_ant.strftime("%d/%m/%Y"),
            "inicio_mes_anterior": primeira_ant.strftime("%d/%m/%Y")
        },
        "qtd": {
            "atual": qtd_atual,
            "ano_anterior": qtd_ant,
            "variacao": ((qtd_atual / qtd_ant) - 1) * 100 if qtd_ant else 0
        },
        "kg": {
            "atual": round(total_kg, 2),
            "ano_anterior": round(total_kg_ant, 2),
            "variacao": ((total_kg / total_kg_ant) - 1) * 100 if total_kg_ant else 0
        },
        "ticket": {
            "atual": round(ticket_atual, 2),
            "ano_anterior": round(ticket_ant, 2),
            "variacao": ((ticket_atual / ticket_ant) - 1) * 100 if ticket_ant else 0
        },
        "preco": {
            "preco_medio_kg": preco_kg,
            "preco_medio_m2": preco_m2,
            "total_kg": round(total_kg, 2),
            "total_m2": round(total_m2, 2),
            "data": ultima.strftime("%d/%m/%Y")
        }
    }

# ======================================================
# SALVAR JSON
# ======================================================
def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

# ======================================================
# EXECUÇÃO
# ======================================================
if __name__ == "__main__":
    df = carregar_excel()
    res = calcular(df)

    salvar("kpi_faturamento.json", res["faturamento"])
    salvar("kpi_quantidade_pedidos.json", res["qtd"])
    salvar("kpi_kg_total.json", res["kg"])
    salvar("kpi_ticket_medio.json", res["ticket"])
    salvar("kpi_preco_medio.json", res["preco"])

    print("\n✓ JSON gerados corretamente.")
    print("📅 Período:", res["faturamento"]["inicio_mes"], "→", res["faturamento"]["data"])
