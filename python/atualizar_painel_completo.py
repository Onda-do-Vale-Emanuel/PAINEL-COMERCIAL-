import pandas as pd
import json
import re
from datetime import datetime
import os

CAMINHO_2026 = "excel/PEDIDOS_2026.xlsx"
CAMINHO_2025 = "excel/PEDIDOS_2025.xlsx"

def limpar_numero(v):

    if pd.isna(v):
        return 0.0

    if isinstance(v, datetime):
        return 0.0

    v = str(v).strip()
    v = re.sub(r"[^0-9,.-]", "", v)

    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0

    if "." in v and "," in v:
        return float(v.replace(".", "").replace(",", "."))

    if "," in v:
        return float(v.replace(",", "."))

    return float(v)

def carregar_planilha(caminho):

    df = pd.read_excel(caminho)
    df.columns = df.columns.str.upper().str.strip()

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    if "TIPO DE PEDIDO" in df.columns:
        df = df[df["TIPO DE PEDIDO"].str.upper() == "NORMAL"]

    for col in ["VALOR COM IPI", "KG", "TOTAL M2"]:
        if col in df.columns:
            df[col] = df[col].apply(limpar_numero)

    return df

def resumo_periodo(df, ano_base):

    hoje = datetime.now()

    inicio = datetime(ano_base, hoje.month, 1)
    fim = datetime(ano_base, hoje.month, hoje.day)

    df_periodo = df[
        (df["DATA"] >= inicio) &
        (df["DATA"] <= fim)
    ]

    total_valor = df_periodo["VALOR COM IPI"].sum()
    total_kg = df_periodo["KG"].sum()
    total_m2 = df_periodo["TOTAL M2"].sum()
    pedidos = len(df_periodo)

    ticket = total_valor / pedidos if pedidos else 0
    preco_kg = total_valor / total_kg if total_kg else 0
    preco_m2 = total_valor / total_m2 if total_m2 else 0

    return {
        "pedidos": pedidos,
        "fat": total_valor,
        "kg": int(total_kg),
        "m2": int(total_m2),
        "ticket": ticket,
        "preco_kg": preco_kg,
        "preco_m2": preco_m2,
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y")
    }

def salvar(nome, dados):

    for pasta in ["dados", "site/dados"]:
        os.makedirs(pasta, exist_ok=True)
        with open(f"{pasta}/{nome}", "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":

    print("=====================================")
    print("Atualizando painel a partir do Excel")
    print("=====================================")

    df_2026 = carregar_planilha(CAMINHO_2026)
    df_2025 = carregar_planilha(CAMINHO_2025)

    atual = resumo_periodo(df_2026, 2026)
    anterior = resumo_periodo(df_2025, 2025)

    salvar("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"] / anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": atual["inicio"],
        "data_atual": atual["fim"],
        "inicio_mes_anterior": anterior["inicio"],
        "data_ano_anterior": anterior["fim"]
    })

    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"] / anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    })

    salvar("kpi_kg_total.json", {
        "atual": atual["kg"],
        "ano_anterior": anterior["kg"],
        "variacao": ((atual["kg"] / anterior["kg"]) - 1) * 100 if anterior["kg"] else 0
    })

    salvar("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": ((atual["ticket"] / anterior["ticket"]) - 1) * 100 if anterior["ticket"] else 0
    })

    salvar("kpi_preco_medio.json", {
        "atual": {
            "preco_medio_kg": round(atual["preco_kg"], 2),
            "preco_medio_m2": round(atual["preco_m2"], 2),
            "data": atual["fim"]
        },
        "ano_anterior": {
            "preco_medio_kg": round(anterior["preco_kg"], 2),
            "preco_medio_m2": round(anterior["preco_m2"], 2),
            "data": anterior["fim"]
        }
    })

    print("\n=====================================")
    print("ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!")
    print("=====================================")
    print("Pedidos 2026:", atual["pedidos"])
    print("Pedidos 2025:", anterior["pedidos"])
