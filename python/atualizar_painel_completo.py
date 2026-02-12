import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ======================================================
# LIMPAR NÚMEROS (BR + DATETIME DO EXCEL)
# ======================================================
def limpar_numero(v):

    if pd.isna(v):
        return 0.0

    # Se vier datetime do Excel (caso do 29/10/1900 01:12:00)
    if isinstance(v, datetime):
        base = datetime(1899, 12, 30)
        dias = (v - base).days
        fracao = (v - base).seconds / 86400
        return round(dias + fracao, 2)

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
# CARREGAR E FILTRAR
# ======================================================
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


# ======================================================
# RESUMO
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
        "preco_m2": preco_m2
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

    print("=====================================")
    print("Atualizando painel a partir do Excel")
    print("=====================================")

    df = carregar()

    hoje = datetime.today()

    # INÍCIO DO MÊS ATUAL
    inicio_atual = hoje.replace(day=1, hour=0, minute=0, second=0)

    # FIM EXIBIÇÃO = HOJE
    fim_atual = hoje

    # ANO ANTERIOR
    inicio_ant = inicio_atual.replace(year=hoje.year - 1)
    fim_ant = fim_atual.replace(year=hoje.year - 1)

    atual = resumo(df, inicio_atual, fim_atual)
    anterior = resumo(df, inicio_ant, fim_ant)

    # ================= FATURAMENTO =================
    salvar("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"]/anterior["fat"])-1)*100 if anterior["fat"] else 0,
        "inicio_mes": inicio_atual.strftime("%d/%m/%Y"),
        "data_atual": fim_atual.strftime("%d/%m/%Y"),
        "inicio_mes_anterior": inicio_ant.strftime("%d/%m/%Y"),
        "data_ano_anterior": fim_ant.strftime("%d/%m/%Y")
    })

    # ================= QUANTIDADE =================
    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"]/anterior["pedidos"])-1)*100 if anterior["pedidos"] else 0
    })

    # ================= KG TOTAL =================
    salvar("kpi_kg_total.json", {
        "atual": round(atual["kg"]),
        "ano_anterior": round(anterior["kg"]),
        "variacao": ((atual["kg"]/anterior["kg"])-1)*100 if anterior["kg"] else 0
    })

    # ================= TICKET =================
    salvar("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": ((atual["ticket"]/anterior["ticket"])-1)*100 if anterior["ticket"] else 0
    })

    # ================= PREÇO MÉDIO =================
    salvar("kpi_preco_medio.json", {
        "atual": {
            "preco_medio_kg": round(atual["preco_kg"], 2),
            "preco_medio_m2": round(atual["preco_m2"], 2),
            "data": fim_atual.strftime("%d/%m/%Y")
        },
        "ano_anterior": {
            "preco_medio_kg": round(anterior["preco_kg"], 2),
            "preco_medio_m2": round(anterior["preco_m2"], 2),
            "data": fim_ant.strftime("%d/%m/%Y")
        }
    })

    print("=====================================")
    print("ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!")
    print("=====================================")
