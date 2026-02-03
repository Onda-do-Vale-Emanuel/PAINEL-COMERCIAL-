import pandas as pd
import json
import re
from datetime import datetime

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
    if "." in v and "," in v:     # 1.234.567,89
        v = v.replace(".", "").replace(",", ".")
        try:
            return float(v)
        except:
            return 0.0
    if "," in v:                 # 123,45
        try:
            return float(v.replace(",", "."))
        except:
            return 0.0
    try:
        return float(v)
    except:
        return 0.0


# ======================================================
# CARREGA PLANILHA / NORMALIZA
# ======================================================
def carregar_excel():
    df = pd.read_excel("excel/PEDIDOS ONDA.xlsx")
    df.columns = df.columns.str.upper().str.strip()

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
# IDENTIFICA PRIMEIRA DATA VÁLIDA E ÚLTIMA DATA DO MÊS
# ======================================================
def obter_intervalo_mes(df):
    ultima = df["DATA"].max()
    if pd.isna(ultima):
        raise Exception("❌ Nenhuma data válida encontrada")

    ano, mes = ultima.year, ultima.month
    df_mes = df[(df["DATA"].dt.year == ano) & (df["DATA"].dt.month == mes)]

    primeira = df_mes["DATA"].min()
    return primeira, ultima  # <= EXATAMENTE O QUE VOCÊ PEDIU


# ======================================================
# CÁLCULO DOS KPIs PADRÃO
# ======================================================
def calcular_kpis_padrao(df, primeira, ultima):
    ano = ultima.year
    mes = ultima.month

    # período atual
    atual = df[(df["DATA"] >= primeira) & (df["DATA"] <= ultima)]

    # período ano anterior
    inicio_a_ant = primeira.replace(year=ano - 1)
    fim_a_ant = ultima.replace(year=ano - 1)
    anterior = df[(df["DATA"] >= inicio_a_ant) & (df["DATA"] <= fim_a_ant)]

    fat_atual = atual["VALOR COM IPI"].sum()
    fat_ant = anterior["VALOR COM IPI"].sum()

    kg_atual = atual["KG"].sum()
    kg_ant = anterior["KG"].sum()

    qtd_atual = len(atual)
    qtd_ant = len(anterior)

    ticket_atual = fat_atual / qtd_atual if qtd_atual else 0
    ticket_ant = fat_ant / qtd_ant if qtd_ant else 0

    return {
        "faturamento": {
            "atual": round(fat_atual, 2),
            "ano_anterior": round(fat_ant, 2),
            "variacao": ((fat_atual / fat_ant - 1) * 100) if fat_ant else 0,
            "data_atual": f"{primeira.strftime('%d/%m/%Y')} até {ultima.strftime('%d/%m/%Y')}",
            "data_ano_anterior": f"{inicio_a_ant.strftime('%d/%m/%Y')} até {fim_a_ant.strftime('%d/%m/%Y')}"
        },
        "kg": {
            "atual": round(kg_atual, 2),
            "ano_anterior": round(kg_ant, 2),
            "variacao": ((kg_atual / kg_ant - 1) * 100) if kg_ant else 0,
        },
        "qtd": {
            "atual": qtd_atual,
            "ano_anterior": qtd_ant,
            "variacao": ((qtd_atual / qtd_ant - 1) * 100) if qtd_ant else 0,
        },
        "ticket": {
            "atual": round(ticket_atual, 2),
            "ano_anterior": round(ticket_ant, 2),
            "variacao": ((ticket_atual / ticket_ant - 1) * 100) if ticket_ant else 0,
        }
    }


# ======================================================
# PREÇO MÉDIO (MESMO PERÍODO)
# ======================================================
def calcular_preco_medio(df, primeira, ultima):
    atual = df[(df["DATA"] >= primeira) & (df["DATA"] <= ultima)]

    total_valor = atual["VALOR COM IPI"].sum()
    total_kg = atual["KG"].sum()
    total_m2 = atual["TOTAL M2"].sum()

    preco_kg = round(total_valor / total_kg, 2) if total_kg else 0
    preco_m2 = round(total_valor / total_m2, 2) if total_m2 else 0

    return {
        "preco_medio_kg": preco_kg,
        "preco_medio_m2": preco_m2,
        "total_kg": round(total_kg, 2),
        "total_m2": round(total_m2, 2),
        "data": ultima.strftime("%d/%m/%Y")
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
# EXECUÇÃO PRINCIPAL
# ======================================================
if __name__ == "__main__":
    df = carregar_excel()

    primeira, ultima = obter_intervalo_mes(df)

    print("📅 Período considerado:", primeira, "→", ultima)

    kpis = calcular_kpis_padrao(df, primeira, ultima)
    preco = calcular_preco_medio(df, primeira, ultima)

    salvar("kpi_faturamento.json", kpis["faturamento"])
    salvar("kpi_kg_total.json", kpis["kg"])
    salvar("kpi_quantidade_pedidos.json", kpis["qtd"])
    salvar("kpi_ticket_medio.json", kpis["ticket"])
    salvar("kpi_preco_medio.json", preco)

    print("\n✓ JSON gerados com sucesso.")
