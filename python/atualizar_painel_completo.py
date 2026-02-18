import pandas as pd
import json
from datetime import datetime
from pathlib import Path

# ============================================
# DEFINIR DIRETÓRIO BASE
# ============================================
BASE_DIR = Path(__file__).resolve().parent.parent

EXCEL_2026 = BASE_DIR / "excel" / "PEDIDOS_2026.xlsx"
EXCEL_2025 = BASE_DIR / "excel" / "PEDIDOS_2025.xlsx"

DADOS_DIR = BASE_DIR / "site" / "dados"

# ============================================
# FUNÇÃO PRINCIPAL
# ============================================
def main(data_inicio=None, data_fim=None):

    # ==========================
    # CARREGAR PLANILHAS
    # ==========================
    df_2026 = pd.read_excel(EXCEL_2026)
    df_2025 = pd.read_excel(EXCEL_2025)

    # NORMALIZAR COLUNAS (ANTI-ERRO)
    df_2026.columns = df_2026.columns.str.strip().str.upper()
    df_2025.columns = df_2025.columns.str.strip().str.upper()

    if "DATA" not in df_2026.columns:
        raise Exception(f"Coluna DATA não encontrada em 2026. Colunas disponíveis: {df_2026.columns.tolist()}")

    if "DATA" not in df_2025.columns:
        raise Exception(f"Coluna DATA não encontrada em 2025. Colunas disponíveis: {df_2025.columns.tolist()}")

    # CONVERTER DATA
    df_2026["DATA"] = pd.to_datetime(df_2026["DATA"], errors="coerce")
    df_2025["DATA"] = pd.to_datetime(df_2025["DATA"], errors="coerce")

    # FILTRAR APENAS NORMAL
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

    # ==========================
    # CÁLCULOS
    # ==========================
    pedidos_atual = len(df_atual)
    pedidos_ant = len(df_ant)

    fat_atual = df_atual["VALOR COM IPI"].sum()
    fat_ant = df_ant["VALOR COM IPI"].sum()

    variacao = ((fat_atual / fat_ant) - 1) * 100 if fat_ant else 0

    faturamento_json = {
        "atual": round(float(fat_atual), 2),
        "ano_anterior": round(float(fat_ant), 2),
        "variacao": variacao,
        "inicio_mes": inicio.strftime("%d/%m/%Y"),
        "data_atual": fim.strftime("%d/%m/%Y"),
        "inicio_mes_anterior": inicio_ant.strftime("%d/%m/%Y"),
        "data_ano_anterior": fim_ant.strftime("%d/%m/%Y")
    }

    quantidade_json = {
        "atual": pedidos_atual,
        "ano_anterior": pedidos_ant,
        "variacao": ((pedidos_atual / pedidos_ant) - 1) * 100 if pedidos_ant else 0
    }

    # ==========================
    # SALVAR JSONS
    # ==========================
    DADOS_DIR.mkdir(parents=True, exist_ok=True)

    with open(DADOS_DIR / "kpi_faturamento.json", "w", encoding="utf-8") as f:
        json.dump(faturamento_json, f, indent=2, ensure_ascii=False)

    with open(DADOS_DIR / "kpi_quantidade_pedidos.json", "w", encoding="utf-8") as f:
        json.dump(quantidade_json, f, indent=2, ensure_ascii=False)

    print("\n=====================================")
    print("ATUALIZAÇÃO CONCLUÍDA COM SUCESSO")
    print("=====================================")
    print("Pedidos 2026:", pedidos_atual)
    print("Pedidos 2025:", pedidos_ant)
    print("Faturamento 2026:", round(fat_atual,2))
    print("Faturamento 2025:", round(fat_ant,2))
    print("=====================================\n")


if __name__ == "__main__":
    main()
