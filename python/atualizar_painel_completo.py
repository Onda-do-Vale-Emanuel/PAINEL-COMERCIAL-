import pandas as pd
import json
from datetime import datetime
from pathlib import Path

# ============================================
# DEFINIR DIRETÓRIO BASE DO PROJETO
# ============================================
BASE_DIR = Path(__file__).resolve().parent.parent

EXCEL_2026 = BASE_DIR / "excel" / "PEDIDOS_2026.xlsx"
EXCEL_2025 = BASE_DIR / "excel" / "PEDIDOS_2025.xlsx"

DADOS_DIR = BASE_DIR / "site" / "dados"

# ============================================
# FUNÇÃO PRINCIPAL
# ============================================
def main(data_inicio=None, data_fim=None):

    df_2026 = pd.read_excel(EXCEL_2026)
    df_2025 = pd.read_excel(EXCEL_2025)

    df_2026["DATA"] = pd.to_datetime(df_2026["DATA"], errors="coerce")
    df_2025["DATA"] = pd.to_datetime(df_2025["DATA"], errors="coerce")

    df_2026 = df_2026[df_2026["TIPO DE PEDIDO"] == "Normal"]
    df_2025 = df_2025[df_2025["TIPO DE PEDIDO"] == "Normal"]

    hoje = datetime.today()

    if data_inicio and data_fim:
        inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
        fim = datetime.strptime(data_fim, "%d/%m/%Y")
    else:
        inicio = hoje.replace(day=1)
        fim = hoje

    inicio_ant = inicio.replace(year=inicio.year - 1)
    fim_ant = fim.replace(year=fim.year - 1)

    df_atual = df_2026[(df_2026["DATA"] >= inicio) & (df_2026["DATA"] <= fim)]
    df_ant = df_2025[(df_2025["DATA"] >= inicio_ant) & (df_2025["DATA"] <= fim_ant)]

    # =============================
    # CÁLCULOS
    # =============================
    pedidos_atual = len(df_atual)
    pedidos_ant = len(df_ant)

    fat_atual = df_atual["VALOR COM IPI"].sum()
    fat_ant = df_ant["VALOR COM IPI"].sum()

    kg_atual = df_atual["KG"].sum()
    kg_ant = df_ant["KG"].sum()

    m2_atual = df_atual["TOTAL M2"].sum()
    m2_ant = df_ant["TOTAL M2"].sum()

    ticket_atual = fat_atual / pedidos_atual if pedidos_atual else 0
    ticket_ant = fat_ant / pedidos_ant if pedidos_ant else 0

    preco_kg_atual = fat_atual / kg_atual if kg_atual else 0
    preco_kg_ant = fat_ant / kg_ant if kg_ant else 0

    preco_m2_atual = fat_atual / m2_atual if m2_atual else 0
    preco_m2_ant = fat_ant / m2_ant if m2_ant else 0

    # =============================
    # JSON FATURAMENTO
    # =============================
    faturamento_json = {
        "atual": round(fat_atual, 2),
        "ano_anterior": round(fat_ant, 2),
        "variacao": ((fat_atual / fat_ant) - 1) * 100 if fat_ant else 0,
        "inicio_mes": inicio.strftime("%d/%m/%Y"),
        "data_atual": fim.strftime("%d/%m/%Y"),
        "inicio_mes_anterior": inicio_ant.strftime("%d/%m/%Y"),
        "data_ano_anterior": fim_ant.strftime("%d/%m/%Y")
    }

    # =============================
    # JSON QUANTIDADE
    # =============================
    quantidade_json = {
        "atual": pedidos_atual,
        "ano_anterior": pedidos_ant,
        "variacao": ((pedidos_atual / pedidos_ant) - 1) * 100 if pedidos_ant else 0
    }

    # =============================
    # JSON KG TOTAL
    # =============================
    kg_json = {
        "atual": round(kg_atual, 0),
        "ano_anterior": round(kg_ant, 0),
        "variacao": ((kg_atual / kg_ant) - 1) * 100 if kg_ant else 0
    }

    # =============================
    # JSON TICKET
    # =============================
    ticket_json = {
        "atual": round(ticket_atual, 2),
        "ano_anterior": round(ticket_ant, 2),
        "variacao": ((ticket_atual / ticket_ant) - 1) * 100 if ticket_ant else 0
    }

    # =============================
    # JSON PREÇO MÉDIO
    # =============================
    preco_json = {
        "atual": {
            "preco_medio_kg": round(preco_kg_atual, 2),
            "preco_medio_m2": round(preco_m2_atual, 2),
            "total_kg": round(kg_atual, 0),
            "total_m2": round(m2_atual, 0),
            "data": fim.strftime("%d/%m/%Y")
        },
        "ano_anterior": {
            "preco_medio_kg": round(preco_kg_ant, 2),
            "preco_medio_m2": round(preco_m2_ant, 2),
            "total_kg": round(kg_ant, 0),
            "total_m2": round(m2_ant, 0),
            "data": fim_ant.strftime("%d/%m/%Y")
        }
    }

    # =============================
    # SALVAR TODOS OS JSONS
    # =============================
    DADOS_DIR.mkdir(exist_ok=True)

    with open(DADOS_DIR / "kpi_faturamento.json", "w", encoding="utf-8") as f:
        json.dump(faturamento_json, f, indent=2, ensure_ascii=False)

    with open(DADOS_DIR / "kpi_quantidade_pedidos.json", "w", encoding="utf-8") as f:
        json.dump(quantidade_json, f, indent=2, ensure_ascii=False)

    with open(DADOS_DIR / "kpi_kg_total.json", "w", encoding="utf-8") as f:
        json.dump(kg_json, f, indent=2, ensure_ascii=False)

    with open(DADOS_DIR / "kpi_ticket_medio.json", "w", encoding="utf-8") as f:
        json.dump(ticket_json, f, indent=2, ensure_ascii=False)

    with open(DADOS_DIR / "kpi_preco_medio.json", "w", encoding="utf-8") as f:
        json.dump(preco_json, f, indent=2, ensure_ascii=False)

    print("ATUALIZAÇÃO CONCLUÍDA")
    print("Pedidos 2026:", pedidos_atual)
    print("Pedidos 2025:", pedidos_ant)


# ============================================
# EXECUÇÃO AUTOMÁTICA
# ============================================
if __name__ == "__main__":
    main()
