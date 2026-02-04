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
# CARREGA EXCEL
# ======================================================
def carregar_excel():
    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.strip().str.upper()

    obrigatorias = ["DATA", "VALOR COM IPI", "KG", "TOTAL M2"]
    for c in obrigatorias:
        if c not in df.columns:
            raise Exception(f"❌ Coluna obrigatória ausente: {c}")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()].copy()

    df["VALOR COM IPI"] = df["VALOR COM IPI"].apply(limpar_numero)
    df["KG"] = df["KG"].apply(limpar_numero)
    df["TOTAL M2"] = df["TOTAL M2"].apply(limpar_numero)

    return df.reset_index(drop=True)

# ======================================================
# PERÍODO CORRETO
# ======================================================
def obter_periodo(df):
    ultima = df["DATA"].max()
    inicio_mes = datetime(ultima.year, ultima.month, 1)
    return inicio_mes, ultima

# ======================================================
# CÁLCULO PRINCIPAL
# ======================================================
def calcular_kpis(df):
    inicio, fim = obter_periodo(df)

    # ===== ATUAL =====
    df_atual = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)].copy()
    df_atual = df_atual.reset_index(drop=True)

    qtd_atual = len(df_atual)
    fat_atual = df_atual["VALOR COM IPI"].sum()
    kg_atual = df_atual["KG"].sum()
    m2_atual = df_atual["TOTAL M2"].sum()

    # ===== ANO ANTERIOR =====
    inicio_ant = inicio.replace(year=inicio.year - 1)
    fim_ant = fim.replace(year=fim.year - 1)

    df_ant = df[(df["DATA"] >= inicio_ant) & (df["DATA"] <= fim_ant)].copy()
    df_ant = df_ant.reset_index(drop=True)

    qtd_ant = len(df_ant)
    fat_ant = df_ant["VALOR COM IPI"].sum()
    kg_ant = df_ant["KG"].sum()
    m2_ant = df_ant["TOTAL M2"].sum()

    # ===== TICKET =====
    ticket_atual = fat_atual / qtd_atual if qtd_atual else 0
    ticket_ant = fat_ant / qtd_ant if qtd_ant else 0

    # ===== PREÇO MÉDIO =====
    preco_kg = fat_atual / kg_atual if kg_atual else 0
    preco_m2 = fat_atual / m2_atual if m2_atual else 0

    print("=====================================")
    print("Pedidos atuais:", qtd_atual)
    print("Período:", inicio.strftime("%d/%m/%Y"), "→", fim.strftime("%d/%m/%Y"))
    print("=====================================")

    return {
        "faturamento": {
            "atual": round(fat_atual, 2),
            "ano_anterior": round(fat_ant, 2),
            "variacao": ((fat_atual / fat_ant) - 1) * 100 if fat_ant else 0,
            "data": fim.strftime("%d/%m/%Y"),
            "inicio_mes": inicio.strftime("%d/%m/%Y"),
            "data_ano_anterior": fim_ant.strftime("%d/%m/%Y"),
            "inicio_mes_anterior": inicio_ant.strftime("%d/%m/%Y"),
        },
        "qtd": {
            "atual": qtd_atual,
            "ano_anterior": qtd_ant,
            "variacao": ((qtd_atual / qtd_ant) - 1) * 100 if qtd_ant else 0
        },
        "kg": {
            "atual": round(kg_atual, 2),
            "ano_anterior": round(kg_ant, 2),
            "variacao": ((kg_atual / kg_ant) - 1) * 100 if kg_ant else 0
        },
        "ticket": {
            "atual": round(ticket_atual, 2),
            "ano_anterior": round(ticket_ant, 2),
            "variacao": ((ticket_atual / ticket_ant) - 1) * 100 if ticket_ant else 0
        },
        "preco": {
            "preco_medio_kg": round(preco_kg, 2),
            "preco_medio_m2": round(preco_m2, 2),
            "total_kg": round(kg_atual, 2),
            "total_m2": round(m2_atual, 2),
            "data": fim.strftime("%d/%m/%Y")
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
    kpis = calcular_kpis(df)

    salvar("kpi_faturamento.json", kpis["faturamento"])
    salvar("kpi_quantidade_pedidos.json", kpis["qtd"])
    salvar("kpi_kg_total.json", kpis["kg"])
    salvar("kpi_ticket_medio.json", kpis["ticket"])
    salvar("kpi_preco_medio.json", kpis["preco"])

    print("✓ JSON gerados corretamente")
