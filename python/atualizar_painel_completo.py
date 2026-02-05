import pandas as pd
import json
import re

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

def limpar_numero(v):
    if pd.isna(v):
        return 0.0
    v = re.sub(r"[^0-9,.-]", "", str(v))
    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0
    if "." in v and "," in v:
        return float(v.replace(".", "").replace(",", "."))
    if "," in v:
        return float(v.replace(",", "."))
    return float(v)

def carregar():
    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.upper().str.strip()

    obrig = ["PEDIDO", "DATA", "VALOR COM IPI", "KG", "TOTAL M2"]
    for c in obrig:
        if c not in df.columns:
            raise Exception(f"Coluna ausente: {c}")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    for c in ["VALOR COM IPI", "KG", "TOTAL M2"]:
        df[c] = df[c].apply(limpar_numero)

    return df

def calcular(df):
    ultima = df["DATA"].max()
    ano = ultima.year
    mes = ultima.month

    df_mes = df[(df["DATA"].dt.year == ano) & (df["DATA"].dt.month == mes)]
    inicio_mes = df_mes["DATA"].min()

    df_periodo = df[(df["DATA"] >= inicio_mes) & (df["DATA"] <= ultima)]

    qtd = df_periodo["PEDIDO"].nunique()
    total = df_periodo["VALOR COM IPI"].sum()
    kg = df_periodo["KG"].sum()
    m2 = df_periodo["TOTAL M2"].sum()

    # ano anterior
    inicio_ant = inicio_mes.replace(year=ano - 1)
    fim_ant = ultima.replace(year=ano - 1)

    df_ant = df[(df["DATA"] >= inicio_ant) & (df["DATA"] <= fim_ant)]

    qtd_ant = df_ant["PEDIDO"].nunique()
    total_ant = df_ant["VALOR COM IPI"].sum()
    kg_ant = df_ant["KG"].sum()

    return {
        "faturamento": {
            "atual": round(total, 2),
            "ano_anterior": round(total_ant, 2),
            "variacao": ((total / total_ant) - 1) * 100 if total_ant else 0,
            "data_atual": ultima.strftime("%d/%m/%Y"),
            "inicio_mes": inicio_mes.strftime("%d/%m/%Y"),
            "data_ano_anterior": fim_ant.strftime("%d/%m/%Y"),
            "inicio_mes_anterior": inicio_ant.strftime("%d/%m/%Y")
        },
        "qtd": {
            "atual": qtd,
            "ano_anterior": qtd_ant,
            "variacao": ((qtd / qtd_ant) - 1) * 100 if qtd_ant else 0
        },
        "kg": {
            "atual": round(kg, 0),
            "ano_anterior": round(kg_ant, 0),
            "variacao": ((kg / kg_ant) - 1) * 100 if kg_ant else 0
        },
        "preco": {
            "preco_medio_kg": round(total / kg, 2) if kg else 0,
            "preco_medio_m2": round(total / m2, 2) if m2 else 0,
            "total_kg": round(kg, 2),
            "total_m2": round(m2, 2),
            "data": ultima.strftime("%d/%m/%Y")
        }
    }

def salvar(nome, dados):
    for p in [f"dados/{nome}", f"site/dados/{nome}"]:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    df = carregar()
    res = calcular(df)

    salvar("kpi_faturamento.json", res["faturamento"])
    salvar("kpi_quantidade_pedidos.json", res["qtd"])
    salvar("kpi_kg_total.json", res["kg"])
    salvar("kpi_preco_medio.json", res["preco"])

    print("Pedidos atuais:", res["qtd"]["atual"])
    print("Período:", res["faturamento"]["inicio_mes"], "→", res["faturamento"]["data_atual"])
