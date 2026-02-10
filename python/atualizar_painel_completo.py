import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

def limpar_numero(v):
    if pd.isna(v):
        return 0.0
    v = str(v).strip()
    v = re.sub(r"[^0-9,.-]", "", v)
    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0
    if "." in v and "," in v:
        return float(v.replace(".", "").replace(",", "."))
    if "," in v:
        return float(v.replace(",", "."))
    try:
        return float(v)
    except:
        return 0.0

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

def obter_periodo(df, ano):
    df_ano = df[df["DATA"].dt.year == ano]
    if len(df_ano) == 0:
        return None, None
    inicio = df_ano["DATA"].min()
    fim = df_ano["DATA"].max()
    return inicio, fim

def resumo(df, inicio, fim):
    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)]
    total_valor = d["VALOR COM IPI"].sum()
    total_kg = d["KG"].sum()
    total_m2 = d["TOTAL M2"].sum()
    pedidos = len(d)
    ticket = total_valor / pedidos if pedidos else 0
    preco_kg = total_valor / total_kg if total_kg else 0
    preco_m2 = total_valor / total_m2 if total_m2 else 0

    return {
        "pedidos": pedidos,
        "fat": total_valor,
        "kg": total_kg,
        "m2": total_m2,
        "ticket": ticket,
        "preco_kg": preco_kg,
        "preco_m2": preco_m2,
        "inicio": inicio.strftime("%d/%m/%Y") if inicio else "--",
        "fim": fim.strftime("%d/%m/%Y") if fim else "--"
    }

def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    df = carregar()

    ano_atual = df["DATA"].max().year
    ano_ant = ano_atual - 1

    inicio_atual, fim_atual = obter_periodo(df, ano_atual)
    inicio_ant, fim_ant = obter_periodo(df, ano_ant)

    atual = resumo(df, inicio_atual, fim_atual)
    anterior = resumo(df, inicio_ant, fim_ant)

    salvar("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"]/anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": atual["inicio"],
        "data_atual": atual["fim"],
        "inicio_mes_anterior": anterior["inicio"],
        "data_ano_anterior": anterior["fim"]
    })

    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"]/anterior["pedidos"]) - 1)*100 if anterior["pedidos"] else 0
    })

    salvar("kpi_kg_total.json", {
        "atual": atual["kg"],
        "ano_anterior": anterior["kg"],
        "variacao": ((atual["kg"]/anterior["kg"]) - 1)*100 if anterior["kg"] else 0
    })

    salvar("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": ((atual["ticket"]/anterior["ticket"]) - 1)*100 if anterior["ticket"] else 0
    })

    salvar("kpi_preco_medio.json", {
        "atual": {
            "preco_medio_kg": round(atual["preco_kg"], 2),
            "preco_medio_m2": round(atual["preco_m2"], 2),
            "total_kg": round(atual["kg"], 2),
            "total_m2": round(atual["m2"], 2),
            "data": atual["fim"]
        },
        "ano_anterior": {
            "preco_medio_kg": round(anterior["preco_kg"], 2),
            "preco_medio_m2": round(anterior["preco_m2"], 2),
            "total_kg": round(anterior["kg"], 2),
            "total_m2": round(anterior["m2"], 2),
            "data": anterior["fim"]
        }
    })

    print("Atualização concluída com sucesso!")
