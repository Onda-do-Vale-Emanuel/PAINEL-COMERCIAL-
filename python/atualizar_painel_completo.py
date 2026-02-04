import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ======================================================
# FUNÇÃO DE LIMPEZA COM PADRÃO BRASILEIRO
# ======================================================
def limpar_numero(valor):
    if pd.isna(valor):
        return 0.0

    v = str(valor).strip()
    v = re.sub(r"[^0-9,.-]", "", v)

    if v in ["", "-", ".", ",", ",-", ".-"]:
        return 0.0

    if "." in v and "," in v:
        return float(v.replace(".", "").replace(",", "."))

    if "," in v:
        return float(v.replace(",", "."))

    return float(v)


# ======================================================
# CARREGAR EXCEL
# ======================================================
def carregar_excel():
    df = pd.read_excel(CAMINHO_EXCEL)

    df.columns = df.columns.str.strip().str.upper()

    obrig = ["DATA", "VALOR COM IPI", "KG", "TOTAL M2"]
    for c in obrig:
        if c not in df.columns:
            raise Exception(f"❌ Coluna obrigatória ausente: {c}")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    df["VALOR COM IPI"] = df["VALOR COM IPI"].apply(limpar_numero)
    df["KG"] = df["KG"].apply(limpar_numero)
    df["TOTAL M2"] = df["TOTAL M2"].apply(limpar_numero)

    return df


# ======================================================
# OBTÉM PERÍODO REAL DO MÊS + DATA PARA EXIBIÇÃO
# ======================================================
def obter_periodo(df):
    ultima_data = df["DATA"].max()
    ano = ultima_data.year
    mes = ultima_data.month

    # PRIMEIRA DATA REAL EXISTENTE NO MÊS
    primeira_real = df[df["DATA"].dt.month == mes]["DATA"].min()

    # DATA PARA EXIBIÇÃO SEMPRE COMEÇA EM 01
    primeira_exibicao = datetime(ano, mes, 1)

    return primeira_real, ultima_data, primeira_exibicao


# ======================================================
# CALCULA KPIs
# ======================================================
def calcular_kpis(df):
    primeira_real, ultima, primeira_exibicao = obter_periodo(df)

    # Período real para cálculo
    df_periodo = df[(df["DATA"] >= primeira_real) & (df["DATA"] <= ultima)]

    total_valor = df_periodo["VALOR COM IPI"].sum()
    total_kg = df_periodo["KG"].sum()
    total_m2 = df_periodo["TOTAL M2"].sum()
    qtd = len(df_periodo)

    # Ano anterior
    ano_ant = ultima.year - 1
    primeira_ant = primeira_real.replace(year=ano_ant)
    ultima_ant = ultima.replace(year=ano_ant)

    df_ant = df[(df["DATA"] >= primeira_ant) & (df["DATA"] <= ultima_ant)]

    fat_ant = df_ant["VALOR COM IPI"].sum()
    kg_ant = df_ant["KG"].sum()
    qtd_ant = len(df_ant)

    ticket_atual = total_valor / qtd if qtd else 0
    ticket_ant = fat_ant / qtd_ant if qtd_ant else 0

    # Preço médio (período real)
    preco_kg = round(total_valor / total_kg, 2) if total_kg else 0
    preco_m2 = round(total_valor / total_m2, 2) if total_m2 else 0

    return {
        "periodo": {
            "primeira_real": primeira_real,
            "ultima": ultima,
            "primeira_exibicao": primeira_exibicao
        },
        "fat": {
            "atual": round(total_valor, 2),
            "ano_anterior": round(fat_ant, 2),
            "variacao": ((total_valor / fat_ant - 1) * 100) if fat_ant else 0,
            "data_atual": ultima.strftime("%d/%m/%Y"),
            "data_ano_anterior": ultima_ant.strftime("%d/%m/%Y"),
            "inicio_exibicao": primeira_exibicao.strftime("%d/%m/%Y")
        },
        "qtd": {
            "atual": qtd,
            "ano_anterior": qtd_ant,
            "variacao": ((qtd / qtd_ant - 1) * 100) if qtd_ant else 0
        },
        "kg": {
            "atual": round(total_kg, 2),
            "ano_anterior": round(kg_ant, 2),
            "variacao": ((total_kg / kg_ant - 1) * 100) if kg_ant else 0
        },
        "ticket": {
            "atual": round(ticket_atual, 2),
            "ano_anterior": round(ticket_ant, 2),
            "variacao": ((ticket_atual / ticket_ant - 1) * 100) if ticket_ant else 0
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
# SALVAR ARQUIVOS JSON
# ======================================================
def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


# ======================================================
# EXECUÇÃO PRINCIPAL
# ======================================================
if __name__ == "__main__":
    df = carregar_excel()
    r = calcular_kpis(df)

    salvar("kpi_faturamento.json", r["fat"])
    salvar("kpi_quantidade_pedidos.json", r["qtd"])
    salvar("kpi_kg_total.json", r["kg"])
    salvar("kpi_ticket_medio.json", r["ticket"])
    salvar("kpi_preco_medio.json", r["preco"])

    print("\n✓ JSON gerados corretamente!")
    print("📅 Exibição no site: 01/" + r["fat"]["data_atual"][3:])
    print("📌 Período real usado:", r["periodo"]["primeira_real"], "→", r["periodo"]["ultima"])
