import pandas as pd
import json
import re
from datetime import datetime

# ======================================================
# 🔥 CAMINHOS DOS ARQUIVOS
# ======================================================
ARQ_2025 = "excel/PEDIDOS_2025.xlsx"
ARQ_2026 = "excel/PEDIDOS_2026.xlsx"

# ======================================================
# 🔥 FUNÇÃO DEFINITIVA PARA LER NÚMEROS BRASILEIROS
# ======================================================
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
    return float(v)

# ======================================================
# 🔥 CARREGAR AS DUAS PLANILHAS E UNIR
# ======================================================
def carregar():
    df25 = pd.read_excel(ARQ_2025)
    df26 = pd.read_excel(ARQ_2026)

    df = pd.concat([df25, df26], ignore_index=True)
    df.columns = df.columns.str.upper().str.strip()

    # limpar colunas numéricas
    colunas_num = ["VALOR TOTAL", "VALOR PRODUTO", "VALOR EMBALAGEM",
                   "VALOR COM IPI", "KG", "TOTAL M2", "PEDIDO"]
    
    for col in colunas_num:
        if col in df.columns:
            df[col] = df[col].apply(limpar_numero)

    # datas
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    # 🔥 manter apenas pedidos válidos (30k–50k)
    df = df[(df["PEDIDO"] >= 30000) & (df["PEDIDO"] <= 50000)]

    return df

# ======================================================
# 🔥 PERÍODOS DOS DOIS ANOS
# ======================================================
def obter_periodos(df):
    ultima = df["DATA"].max()

    inicio_atual = ultima.replace(day=1)

    ano_ant = ultima.year - 1
    inicio_ant = inicio_atual.replace(year=ano_ant)

    fim_atual = ultima

    df_ant = df[(df["DATA"].dt.year == ano_ant) &
                (df["DATA"].dt.month == ultima.month)]

    fim_ant = df_ant["DATA"].max() if len(df_ant) else inicio_ant

    return (inicio_atual, fim_atual), (inicio_ant, fim_ant)

# ======================================================
# 🔥 RESUMO DE QUALQUER PERÍODO
# ======================================================
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
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y")
    }

# ======================================================
# 🔥 SALVAR JSON EM DUPLICIDADE (PASTA raiz E site/)
# ======================================================
def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

# ======================================================
# 🔥 EXECUÇÃO PRINCIPAL
# ======================================================
if __name__ == "__main__":
    df = carregar()
    (inicio_atual, fim_atual), (inicio_ant, fim_ant) = obter_periodos(df)

    atual = resumo(df, inicio_atual, fim_atual)
    anterior = resumo(df, inicio_ant, fim_ant)

    # FATURAMENTO
    salvar("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"]/anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": atual["inicio"],
        "data_atual": atual["fim"],
        "inicio_mes_anterior": anterior["inicio"],
        "data_ano_anterior": anterior["fim"]
    })

    # QTD PEDIDOS
    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"]/anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    })

    # KG TOTAL
    salvar("kpi_kg_total.json", {
        "atual": atual["kg"],
        "ano_anterior": anterior["kg"],
        "variacao": ((atual["kg"]/anterior["kg"]) - 1) * 100 if anterior["kg"] else 0
    })

    # TICKET MÉDIO
    salvar("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": ((atual["ticket"]/anterior["ticket"]) - 1) * 100 if anterior["ticket"] else 0
    })

    # PREÇO MÉDIO
    salvar("kpi_preco_medio.json", {
        "preco_medio_kg": round(atual["preco_kg"], 2),
        "preco_medio_m2": round(atual["preco_m2"], 2),
        "total_kg": round(atual["kg"], 2),
        "total_m2": round(atual["m2"], 2),
        "data": atual["fim"]
    })

    # PREÇO MÉDIO ANO ANTERIOR (NOVO)
    salvar("kpi_preco_medio_ano_anterior.json", {
        "preco_medio_kg": round(anterior["preco_kg"], 2),
        "preco_medio_m2": round(anterior["preco_m2"], 2),
        "total_kg": round(anterior["kg"], 2),
        "total_m2": round(anterior["m2"], 2),
        "data": anterior["fim"]
    })

    print("=====================================")
    print("PAINEL ATUALIZADO COM SUCESSO!")
    print("=====================================")
