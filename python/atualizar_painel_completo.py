import pandas as pd
import json
import os
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_EXCEL = os.path.join(BASE_DIR, "..", "excel", "PEDIDOS ONDA.xlsx")
CAMINHO_DADOS = os.path.join(BASE_DIR, "..", "dados")
CAMINHO_SITE = os.path.join(BASE_DIR, "..", "site", "dados")

# ============================================================
# FUNÇÃO CORRIGIDA PARA CARREGAR O EXCEL (ACEITA QUALQUER "DATA")
# ============================================================
def carregar_excel():
    df = pd.read_excel(CAMINHO_EXCEL)

    # Normalizar nomes
    df.columns = df.columns.str.strip().str.upper()

    # Procurar qualquer coluna que contenha "DATA"
    col_data = None
    for col in df.columns:
        if "DATA" in col:
            col_data = col
            break

    if not col_data:
        raise Exception("❌ A coluna 'DATA' não foi encontrada no Excel.")

    # Converter para datetime
    df[col_data] = pd.to_datetime(df[col_data], errors="coerce")
    df = df.dropna(subset=[col_data])

    # Padronizar nome interno
    df = df.rename(columns={col_data: "DATA"})

    return df


# ============================================================
# ENCONTRAR A ÚLTIMA DATA VÁLIDA DO MÊS
# ============================================================
def ultima_data_valida(df):
    hoje = datetime.now().date()

    # ① Primeiro tenta pegar pedidos do dia
    df_hoje = df[df["DATA"].dt.date == hoje]
    if len(df_hoje) > 0:
        return hoje

    # ② Depois o dia imediatamente anterior
    for i in range(1, 10):
        dia = hoje - timedelta(days=i)
        df_dia = df[df["DATA"].dt.date == dia]
        if len(df_dia) > 0:
            return dia

    # ③ Caso seja início de mês e não existam pedidos, pega o PRIMEIRO dia com pedido no mês
    mes_atual = hoje.month
    ano_atual = hoje.year
    df_mes = df[(df["DATA"].dt.month == mes_atual) & (df["DATA"].dt.year == ano_atual)]

    if len(df_mes) > 0:
        return df_mes["DATA"].dt.date.min()

    # ④ Último recurso
    return df["DATA"].dt.date.max()


# ============================================================
# CALCULAR KPIS PADRÃO
# ============================================================
def calcular_kpis_padrao(df, data_ref):
    df_atual = df[df["DATA"].dt.date == data_ref]

    data_ano_anterior = data_ref.replace(year=data_ref.year - 1)
    df_anterior = df[df["DATA"].dt.date == data_ano_anterior]

    def soma(df):
        return df["VALOR TOTAL C/ IPI"].sum()

    def kg(df):
        return df["KG"].sum()

    atual = soma(df_atual)
    anterior = soma(df_anterior)

    kg_atual = kg(df_atual)
    kg_anterior = kg(df_anterior)

    qtd_atual = len(df_atual)
    qtd_anterior = len(df_anterior)

    def variacao(a, b):
        if b == 0:
            return 0
        return ((a - b) / b) * 100

    return {
        "faturamento": {
            "atual": round(atual, 2),
            "ano_anterior": round(anterior, 2),
            "variacao": round(variacao(atual, anterior), 2),
            "data_atual": data_ref.strftime("%d/%m/%Y"),
            "data_ano_anterior": data_ano_anterior.strftime("%d/%m/%Y"),
        },
        "kg_total": {
            "atual": round(kg_atual, 2),
            "ano_anterior": round(kg_anterior, 2),
            "variacao": round(variacao(kg_atual, kg_anterior), 2),
        },
        "quantidade": {
            "atual": qtd_atual,
            "ano_anterior": qtd_anterior,
            "variacao": round(variacao(qtd_atual, qtd_anterior), 2),
        },
        "ticket": {
            "atual": round(atual / qtd_atual, 2) if qtd_atual else 0,
            "ano_anterior": round(anterior / qtd_anterior, 2) if qtd_anterior else 0,
            "variacao": round(
                variacao(
                    round(atual / qtd_atual, 2) if qtd_atual else 0,
                    round(anterior / qtd_anterior, 2) if qtd_anterior else 0,
                ),
                2,
            ),
        },
    }


# ============================================================
# GERAR JSON INDIVIDUAL
# ============================================================
def salvar_json(nome, dados):
    with open(os.path.join(CAMINHO_DADOS, nome), "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

    with open(os.path.join(CAMINHO_SITE, nome), "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


# ============================================================
# PREÇO MÉDIO KG / M2
# ============================================================
def salvar_preco_medio(df):
    total_kg = df["KG"].sum()
    total_m2 = df["M2"].sum() if "M2" in df.columns else 0
    total_valor = df["VALOR TOTAL C/ IPI"].sum()

    preco_kg = total_valor / total_kg if total_kg else 0
    preco_m2 = total_valor / total_m2 if total_m2 else 0

    dados = {
        "preco_medio_kg": round(preco_kg, 2),
        "preco_medio_m2": round(preco_m2, 2),
        "total_kg": round(total_kg, 2),
        "total_m2": round(total_m2, 2),
        "data": datetime.now().strftime("%d/%m/%Y"),
    }

    salvar_json("kpi_preco_medio.json", dados)

    print("Preço médio gerado:")
    print(json.dumps(dados, indent=2, ensure_ascii=False))


# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================
def calcular_kpis():
    df = carregar_excel()
    data_ref = ultima_data_valida(df)

    print("📅 Última data encontrada no Excel:", data_ref)

    kpis = calcular_kpis_padrao(df, data_ref)

    salvar_json("kpi_faturamento.json", {
        **kpis["faturamento"],
        "atual_qtd": kpis["quantidade"]["atual"],
        "ano_anterior_qtd": kpis["quantidade"]["ano_anterior"],
    })

    salvar_json("kpi_quantidade_pedidos.json", kpis["quantidade"])
    salvar_json("kpi_ticket_medio.json", kpis["ticket"])
    salvar_json("kpi_kg_total.json", kpis["kg_total"])

    salvar_preco_medio(df)


# ============================================================
# EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    calcular_kpis()
    print("\nKPIs gerados com sucesso!\n")
