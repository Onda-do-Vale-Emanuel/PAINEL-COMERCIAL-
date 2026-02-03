import pandas as pd
import json
import re
from datetime import datetime

# ======================================================
# 🔥 FUNÇÃO DEFINITIVA PARA TRATAR NÚMEROS BRASILEIROS
# ======================================================
def limpar_numero(valor):
    if pd.isna(valor):
        return 0.0

    v = str(valor).strip()
    v = re.sub(r"[^0-9,.-]", "", v)

    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0

    # Possui milhares + vírgula decimal → "1.234.567,89"
    if "." in v and "," in v:
        v = v.replace(".", "").replace(",", ".")
        try: return float(v)
        except: return 0.0

    # Apenas vírgula decimal → "123,45"
    if "," in v:
        v = v.replace(",", ".")
        try: return float(v)
        except: return 0.0

    # Apenas ponto:
    partes = v.split(".")
    if len(partes[-1]) == 3:   # "1.234" → milhar
        v = v.replace(".", "")
        try: return float(v)
        except: return 0.0

    # Caso seja decimal
    try: return float(v)
    except: return 0.0


# ======================================================
# 📌 CARREGAR EXCEL
# ======================================================
def carregar_excel():
    df = pd.read_excel("excel/PEDIDOS ONDA.xlsx")
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
# 📌 OBTÉM A ÚLTIMA DATA REAL DO EXCEL
# ======================================================
def obter_data_ref(df):
    datas = df["DATA"].dropna()
    return datas.max()


# ======================================================
# 📌 OBTÉM A PRIMEIRA DATA REAL DO MESMO MÊS
# ======================================================
def obter_primeira_data_mes(df, ano, mes):
    df_mes = df[(df["DATA"].dt.year == ano) & (df["DATA"].dt.month == mes)]
    if df_mes.empty:
        return None
    return df_mes["DATA"].min()


# ======================================================
# 📌 CALCULA KPIs DO PERÍODO REAL (primeira ↔ última)
# ======================================================
def calcular_kpis(df, data_inicio, data_fim):
    df_periodo = df[(df["DATA"] >= data_inicio) & (df["DATA"] <= data_fim)]

    # FATURAMENTO
    fat_atual = df_periodo["VALOR COM IPI"].sum()

    # KG
    kg_atual = df_periodo["KG"].sum()

    # QUANTIDADE
    qtd_atual = len(df_periodo)

    # TICKET
    ticket_atual = fat_atual / qtd_atual if qtd_atual > 0 else 0

    # ANO ANTERIOR
    ano_ant = data_fim.year - 1
    df_ant = df[
        (df["DATA"].dt.year == ano_ant) &
        (df["DATA"].dt.month == data_fim.month) &
        (df["DATA"].dt.day >= data_inicio.day) &
        (df["DATA"].dt.day <= data_fim.day)
    ]

    fat_ant = df_ant["VALOR COM IPI"].sum()
    kg_ant = df_ant["KG"].sum()
    qtd_ant = len(df_ant)
    ticket_ant = fat_ant / qtd_ant if qtd_ant > 0 else 0

    return {
        "faturamento": {
            "atual": round(fat_atual, 2),
            "ano_anterior": round(fat_ant, 2),
            "variacao": ((fat_atual / fat_ant - 1) * 100) if fat_ant > 0 else 0,
            "data_atual": data_fim.strftime("%d/%m/%Y"),
            "data_ano_anterior": data_fim.replace(year=ano_ant).strftime("%d/%m/%Y"),
            "primeiro_dia": data_inicio.strftime("%d/%m/%Y"),
        },
        "kg": {
            "atual": round(kg_atual, 2),
            "ano_anterior": round(kg_ant, 2),
            "variacao": ((kg_atual / kg_ant - 1) * 100) if kg_ant > 0 else 0,
        },
        "qtd": {
            "atual": qtd_atual,
            "ano_anterior": qtd_ant,
            "variacao": ((qtd_atual / qtd_ant - 1) * 100) if qtd_ant > 0 else 0,
        },
        "ticket": {
            "atual": round(ticket_atual, 2),
            "ano_anterior": round(ticket_ant, 2),
            "variacao": ((ticket_atual / ticket_ant - 1) * 100) if ticket_ant > 0 else 0,
        }
    }


# ======================================================
# 📌 PREÇO MÉDIO — MESMO PERÍODO REAL
# ======================================================
def calcular_preco_medio(df, data_inicio, data_fim):
    df_periodo = df[(df["DATA"] >= data_inicio) & (df["DATA"] <= data_fim)]

    total_valor = df_periodo["VALOR COM IPI"].sum()
    total_kg = df_periodo["KG"].sum()
    total_m2 = df_periodo["TOTAL M2"].sum()

    preco_kg = round(total_valor / total_kg, 2) if total_kg > 0 else 0
    preco_m2 = round(total_valor / total_m2, 2) if total_m2 > 0 else 0

    return {
        "preco_medio_kg": preco_kg,
        "preco_medio_m2": preco_m2,
        "total_kg": round(total_kg, 2),
        "total_m2": round(total_m2, 2),
        "data": data_fim.strftime("%d/%m/%Y"),
        "primeiro_dia": data_inicio.strftime("%d/%m/%Y")
    }


# ======================================================
# 📌 SALVAR JSON
# ======================================================
def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)

    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)


# ======================================================
# 🚀 EXECUÇÃO PRINCIPAL
# ======================================================
if __name__ == "__main__":
    df = carregar_excel()
    data_fim = obter_data_ref(df)

    ano = data_fim.year
    mes = data_fim.month
    data_inicio = obter_primeira_data_mes(df, ano, mes)

    print("Período usado:", data_inicio, "→", data_fim)

    kpis = calcular_kpis(df, data_inicio, data_fim)
    preco = calcular_preco_medio(df, data_inicio, data_fim)

    salvar("kpi_faturamento.json", kpis["faturamento"])
    salvar("kpi_kg_total.json", kpis["kg"])
    salvar("kpi_quantidade_pedidos.json", kpis["qtd"])
    salvar("kpi_ticket_medio.json", kpis["ticket"])
    salvar("kpi_preco_medio.json", preco)

    print("✓ JSON atualizados com sucesso.")
