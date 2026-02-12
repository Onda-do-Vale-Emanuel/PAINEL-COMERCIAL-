import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ======================================================
# FUNÇÃO PARA LIMPAR NÚMEROS (BR)
# ======================================================
def limpar_numero(v):
    if pd.isna(v):
        return 0.0

    if isinstance(v, (int, float)):
        return float(v)

    v = str(v).strip()
    v = re.sub(r"[^\d,.\-]", "", v)

    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0

    if "." in v and "," in v:
        return float(v.replace(".", "").replace(",", "."))

    if "," in v:
        return float(v.replace(",", "."))

    return float(v)


# ======================================================
# CARREGAR E FILTRAR PLANILHA
# ======================================================
def carregar():
    df = pd.read_excel(CAMINHO_EXCEL)

    df.columns = df.columns.str.upper().str.strip()

    # Datas
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    # Tipo pedido = Normal
    if "TIPO DE PEDIDO" in df.columns:
        df = df[df["TIPO DE PEDIDO"].str.strip() == "Normal"]

    # Limpar colunas numéricas
    for col in ["VALOR COM IPI", "KG", "TOTAL M2"]:
        if col in df.columns:
            df[col] = df[col].apply(limpar_numero)

    # Pedido como string para evitar perda por formato
    df["PEDIDO"] = df["PEDIDO"].astype(str)

    return df


# ======================================================
# PERÍODO CORRETO (USANDO ÚLTIMA DATA REAL)
# ======================================================
def obter_periodos(df):
    ultima_data = df["DATA"].max()

    inicio_atual = ultima_data.replace(day=1)
    fim_atual = ultima_data

    ano_ant = ultima_data.year - 1
    inicio_ant = inicio_atual.replace(year=ano_ant)
    fim_ant = fim_atual.replace(year=ano_ant)

    return (inicio_atual, fim_atual), (inicio_ant, fim_ant)


# ======================================================
# RESUMO
# ======================================================
def resumo(df, inicio, fim):
    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)]

    total_valor = d["VALOR COM IPI"].sum()
    total_kg = d["KG"].sum()
    total_m2 = d["TOTAL M2"].sum()

    pedidos_unicos = d["PEDIDO"].nunique()

    ticket = total_valor / pedidos_unicos if pedidos_unicos else 0
    preco_kg = total_valor / total_kg if total_kg else 0
    preco_m2 = total_valor / total_m2 if total_m2 else 0

    return {
        "pedidos": int(pedidos_unicos),
        "fat": round(total_valor, 2),
        "kg": round(total_kg, 0),
        "m2": round(total_m2, 0),
        "ticket": round(ticket, 2),
        "preco_kg": round(preco_kg, 2),
        "preco_m2": round(preco_m2, 2),
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y"),
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
    print("=====================================")
    print("Atualizando painel a partir do Excel")
    print("=====================================")

    df = carregar()

    (inicio_atual, fim_atual), (inicio_ant, fim_ant) = obter_periodos(df)

    atual = resumo(df, inicio_atual, fim_atual)
    anterior = resumo(df, inicio_ant, fim_ant)

    # FATURAMENTO
    salvar("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"] / anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": atual["inicio"],
        "data_atual": atual["fim"],
        "inicio_mes_anterior": anterior["inicio"],
        "data_ano_anterior": anterior["fim"]
    })

    # QUANTIDADE
    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"] / anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    })

    # KG
    salvar("kpi_kg_total.json", {
        "atual": atual["kg"],
        "ano_anterior": anterior["kg"],
        "variacao": ((atual["kg"] / anterior["kg"]) - 1) * 100 if anterior["kg"] else 0
    })

    # TICKET
    salvar("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": ((atual["ticket"] / anterior["ticket"]) - 1) * 100 if anterior["ticket"] else 0
    })

    # PREÇO MÉDIO
    salvar("kpi_preco_medio.json", {
        "preco_medio_kg": atual["preco_kg"],
        "preco_medio_m2": atual["preco_m2"],
        "total_kg": atual["kg"],
        "total_m2": atual["m2"],
        "data": atual["fim"]
    })

    print("=====================================")
    print("ATUALIZAÇÃO CONCLUÍDA COM SUCESSO")
    print("=====================================")
