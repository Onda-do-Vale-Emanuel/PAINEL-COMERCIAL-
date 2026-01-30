import pandas as pd
import json
import datetime
import os

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"
PASTA_DADOS = "dados"
PASTA_SITE = "site/dados"

def salvar_json(nome, conteudo):
    with open(os.path.join(PASTA_DADOS, nome), "w", encoding="utf-8") as f:
        json.dump(conteudo, f, ensure_ascii=False, indent=2)

    with open(os.path.join(PASTA_SITE, nome), "w", encoding="utf-8") as f:
        json.dump(conteudo, f, ensure_ascii=False, indent=2)


def carregar_excel():
    df = pd.read_excel(CAMINHO_EXCEL)

    df.columns = df.columns.str.strip()

    if "DATA" not in df.columns:
        raise Exception("❌ A coluna 'DATA' não foi encontrada no Excel.")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")

    df = df.dropna(subset=["DATA"])

    return df


def obter_ultima_data_valida(df):
    """Retorna a MAIOR data disponível no Excel (data mais recente)."""
    return df["DATA"].max().date()


def obter_primeira_data_valida_do_mes(df, data_final):
    """
    Retorna a MENOR data existente no Excel dentro do mês/ano da data_final.
    Exemplo:
      Excel tem registros em:
         03/01/2026
         04/01/2026
         08/01/2026
      → retorno = 03/01/2026
    """

    mes = data_final.month
    ano = data_final.year

    df_mes = df[(df["DATA"].dt.month == mes) & (df["DATA"].dt.year == ano)]

    if df_mes.empty:
        # fallback para 01/mes (situação rara, mas segura)
        return datetime.date(ano, mes, 1)

    return df_mes["DATA"].min().date()


def filtrar_periodo(df, data_inicio_real, data_final):
    """Filtra do primeiro dia REAL do mês até a data final."""
    return df[(df["DATA"] >= pd.Timestamp(data_inicio_real)) &
              (df["DATA"] <= pd.Timestamp(data_final))]


def calcular_kpis():
    df = carregar_excel()

    # 🔥 PEGA A ÚLTIMA DATA REAL DO EXCEL
    data_atual = obter_ultima_data_valida(df)

    # 🔥 PEGA O PRIMEIRO DIA REAL COM PEDIDO NO MÊS
    data_inicio_mes = obter_primeira_data_valida_do_mes(df, data_atual)

    # 🟧 ANO ANTERIOR — PARÂMETROS EQUIVALENTES
    try:
        data_ano_anterior = data_atual.replace(year=data_atual.year - 1)
    except:
        data_ano_anterior = data_atual - datetime.timedelta(days=365)

    data_inicio_mes_ant = data_inicio_mes.replace(year=data_inicio_mes.year - 1)

    df_atual = filtrar_periodo(df, data_inicio_mes, data_atual)
    df_ano_anterior = filtrar_periodo(df, data_inicio_mes_ant, data_ano_anterior)

    # Garantir colunas importantes
    for col in ["VALOR_TOTAL_IPI", "KG", "M2"]:
        if col not in df.columns:
            df[col] = 0

    # ============================
    # QUANTIDADE DE PEDIDOS
    # ============================
    kpi_qtd = {
        "atual": df_atual["PEDIDO"].nunique(),
        "ano_anterior": df_ano_anterior["PEDIDO"].nunique(),
    }
    kpi_qtd["variacao"] = round(
        ((kpi_qtd["atual"] - kpi_qtd["ano_anterior"]) / max(kpi_qtd["ano_anterior"], 1)) * 100, 1
    )
    salvar_json("kpi_quantidade_pedidos.json", kpi_qtd)

    # ============================
    # FATURAMENTO
    # ============================
    fat_atual = df_atual["VALOR_TOTAL_IPI"].sum()
    fat_ant = df_ano_anterior["VALOR_TOTAL_IPI"].sum()

    kpi_fat = {
        "atual": round(fat_atual, 2),
        "ano_anterior": round(fat_ant, 2),
        "variacao": round(((fat_atual - fat_ant) / max(fat_ant, 1)) * 100, 1),
        "data_atual": data_atual.strftime("%d/%m/%Y"),
        "data_ano_anterior": data_ano_anterior.strftime("%d/%m/%Y"),
    }
    salvar_json("kpi_faturamento.json", kpi_fat)

    # ============================
    # KG TOTAL
    # ============================
    kg_atual = df_atual["KG"].sum()
    kg_ant = df_ano_anterior["KG"].sum()

    kpi_kg = {
        "atual": round(kg_atual, 2),
        "ano_anterior": round(kg_ant, 2),
        "variacao": round(((kg_atual - kg_ant) / max(kg_ant, 1)) * 100, 1),
        "data": data_atual.strftime("%d/%m/%Y"),
    }
    salvar_json("kpi_kg_total.json", kpi_kg)

    # ============================
    # TICKET MÉDIO
    # ============================
    ticket_atual = fat_atual / max(kpi_qtd["atual"], 1)
    ticket_ant = fat_ant / max(kpi_qtd["ano_anterior"], 1)

    kpi_ticket = {
        "atual": round(ticket_atual, 2),
        "ano_anterior": round(ticket_ant, 2),
        "variacao": round(((ticket_atual - ticket_ant) / max(ticket_ant, 1)) * 100, 1),
    }
    salvar_json("kpi_ticket_medio.json", kpi_ticket)

    # ============================
    # PREÇO MÉDIO (NOVO SLIDE)
    # ============================
    total_kg = df_atual["KG"].sum()
    total_m2 = df_atual["M2"].sum()

    preco_medio_kg = fat_atual / max(total_kg, 1)
    preco_medio_m2 = fat_atual / max(total_m2, 1)

    kpi_preco = {
        "preco_medio_kg": round(preco_medio_kg, 2),
        "preco_medio_m2": round(preco_medio_m2, 2),
        "total_kg": round(total_kg, 2),
        "total_m2": round(total_m2, 2),
        "data": data_atual.strftime("%d/%m/%Y"),
    }
    salvar_json("kpi_preco_medio.json", kpi_preco)

    print("KPIs gerados com sucesso!")


if __name__ == "__main__":
    calcular_kpis()
