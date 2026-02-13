import pandas as pd
import json
from datetime import datetime

CAMINHO_2025 = "excel/PEDIDOS_2025.xlsx"
CAMINHO_2026 = "excel/PEDIDOS_2026.xlsx"

# ==========================================
# LIMPEZA SEGURA DE NÚMEROS (VERSÃO FINAL)
# ==========================================
def limpar_numero(v):

    if pd.isna(v):
        return 0.0

    # Se já for número real, retorna direto
    if isinstance(v, (int, float)):
        return float(v)

    v = str(v).strip()

    # Remove símbolo moeda
    v = v.replace("R$", "").strip()

    # Se formato brasileiro 1.234.567,89
    if "," in v and "." in v:
        v = v.replace(".", "").replace(",", ".")

    # Se só vírgula decimal
    elif "," in v:
        v = v.replace(",", ".")

    try:
        return float(v)
    except:
        return 0.0


# ==========================================
# CARREGAR PLANILHA
# ==========================================
def carregar_planilha(caminho):

    df = pd.read_excel(caminho)
    df.columns = df.columns.str.upper().str.strip()

    # Converter DATA
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    hoje = datetime.now()

    # Filtrar Fevereiro até hoje
    df = df[
        (df["DATA"].dt.month == 2) &
        (df["DATA"].dt.day <= hoje.day)
    ]

    # Filtrar apenas Normal
    if "TIPO DE PEDIDO" in df.columns:
        df = df[df["TIPO DE PEDIDO"].astype(str).str.upper() == "NORMAL"]

    # Limpar colunas numéricas
    for col in ["VALOR COM IPI", "KG", "TOTAL M2"]:
        if col in df.columns:
            df[col] = df[col].apply(limpar_numero)

    return df


# ==========================================
# RESUMO
# ==========================================
def resumo(df):

    pedidos = len(df)
    total_valor = df["VALOR COM IPI"].sum()
    total_kg = df["KG"].sum()
    total_m2 = df["TOTAL M2"].sum()

    ticket = total_valor / pedidos if pedidos else 0
    preco_kg = total_valor / total_kg if total_kg else 0
    preco_m2 = total_valor / total_m2 if total_m2 else 0

    return {
        "pedidos": pedidos,
        "fat": round(total_valor, 2),
        "kg": round(total_kg, 0),
        "m2": round(total_m2, 0),
        "ticket": round(ticket, 2),
        "preco_kg": round(preco_kg, 2),
        "preco_m2": round(preco_m2, 2)
    }


# ==========================================
# SALVAR JSON
# ==========================================
def salvar(nome, dados):

    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)

    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)


# ==========================================
# EXECUÇÃO
# ==========================================
if __name__ == "__main__":

    print("=====================================")
    print("Atualizando painel corretamente")
    print("=====================================")

    df_2026 = carregar_planilha(CAMINHO_2026)
    df_2025 = carregar_planilha(CAMINHO_2025)

    atual = resumo(df_2026)
    anterior = resumo(df_2025)

    hoje_str = datetime.now().strftime("%d/%m/%Y")

    # FATURAMENTO
    salvar("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"] / anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": "01/02/2026",
        "data_atual": hoje_str,
        "inicio_mes_anterior": "01/02/2025",
        "data_ano_anterior": hoje_str
    })

    # QUANTIDADE
    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"] / anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    })

    # KG TOTAL
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
        "atual": {
            "preco_medio_kg": atual["preco_kg"],
            "preco_medio_m2": atual["preco_m2"],
            "data": hoje_str
        },
        "ano_anterior": {
            "preco_medio_kg": anterior["preco_kg"],
            "preco_medio_m2": anterior["preco_m2"],
            "data": hoje_str
        }
    })

    print("=====================================")
    print(" ATUALIZAÇÃO FINAL CONCLUÍDA!")
    print("=====================================")
    print("Pedidos 2026:", atual["pedidos"])
    print("Pedidos 2025:", anterior["pedidos"])
    print("=====================================")
