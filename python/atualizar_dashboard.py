import pandas as pd
import json
from datetime import datetime
from pathlib import Path
import numpy as np

print("🔄 Iniciando atualização dos KPIs...")

# ======================================================
# CAMINHOS DO PROJETO
# ======================================================
BASE_DIR = Path(__file__).resolve().parents[1]
ARQ_EXCEL = BASE_DIR / "excel" / "PEDIDOS ONDA.xlsx"
PASTA_DADOS = BASE_DIR / "site" / "dados"
PASTA_DADOS.mkdir(exist_ok=True)

# ======================================================
# ÍNDICES DAS COLUNAS (base 0)
# ======================================================
COL_TIPO = 3          # Tipo de pedido
COL_DATA = 4          # Data
COL_VALOR = 7         # Valor Total
COL_PEDIDO = 1        # Número do pedido

# ======================================================
# DATA DE REFERÊNCIA
# ======================================================
HOJE = datetime(2026, 1, 26)
ANO_ATUAL = HOJE.year
ANO_ANTERIOR = HOJE.year - 1

# ======================================================
# LEITURA DA PLANILHA
# ======================================================
df = pd.read_excel(ARQ_EXCEL)
print(f"📄 Linhas lidas: {len(df)}")

# ======================================================
# FILTRAR PEDIDOS NORMAL
# ======================================================
df = df[df.iloc[:, COL_TIPO].astype(str).str.upper().str.strip() == "NORMAL"]
print(f"✅ NORMAL: {len(df)} linhas")

# ======================================================
# TRATAR DATA
# ======================================================
df["DATA_OK"] = pd.to_datetime(
    df.iloc[:, COL_DATA],
    errors="coerce",
    dayfirst=True
)
df = df.dropna(subset=["DATA_OK"])

# ======================================================
# TRATAR VALOR
# ======================================================
df["VALOR_OK"] = (
    df.iloc[:, COL_VALOR]
    .astype(str)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
)
df["VALOR_OK"] = pd.to_numeric(df["VALOR_OK"], errors="coerce").fillna(0)

# ======================================================
# FILTRAR PERÍODOS
# ======================================================
df_atual = df[
    (df["DATA_OK"].dt.year == ANO_ATUAL) &
    (df["DATA_OK"] <= HOJE)
]

df_ano_anterior = df[
    (df["DATA_OK"].dt.year == ANO_ANTERIOR) &
    (df["DATA_OK"] <= HOJE.replace(year=ANO_ANTERIOR))
]

print(f"📆 Linhas ano atual: {len(df_atual)}")
print(f"📆 Linhas ano anterior: {len(df_ano_anterior)}")

# ======================================================
# KPI 1 — FATURAMENTO
# ======================================================
fat_atual = float(df_atual["VALOR_OK"].sum())
fat_ano_anterior = float(df_ano_anterior["VALOR_OK"].sum())

if fat_ano_anterior > 0:
    var_faturamento = ((fat_atual - fat_ano_anterior) / fat_ano_anterior) * 100
else:
    var_faturamento = None

# ======================================================
# KPI 2 — QUANTIDADE DE PEDIDOS
# (1 linha = 1 pedido, igual Excel)
# ======================================================
qtd_atual = int(len(df_atual))
qtd_ano_anterior = int(len(df_ano_anterior))

if qtd_ano_anterior > 0:
    var_qtd = ((qtd_atual - qtd_ano_anterior) / qtd_ano_anterior) * 100
else:
    var_qtd = None

# ======================================================
# KPI 3 — TICKET MÉDIO
# ======================================================
ticket_atual = fat_atual / qtd_atual if qtd_atual > 0 else 0
ticket_ano_anterior = (
    fat_ano_anterior / qtd_ano_anterior if qtd_ano_anterior > 0 else 0
)

if ticket_ano_anterior > 0:
    var_ticket = ((ticket_atual - ticket_ano_anterior) / ticket_ano_anterior) * 100
else:
    var_ticket = None

# ======================================================
# GERAR JSONs
# ======================================================

# FATURAMENTO
with open(PASTA_DADOS / "kpi_faturamento.json", "w", encoding="utf-8") as f:
    json.dump({
        "atual": round(fat_atual, 2),
        "ano_anterior": round(fat_ano_anterior, 2),
        "variacao": round(var_faturamento, 1) if var_faturamento is not None else None
    }, f, ensure_ascii=False, indent=2)

# QUANTIDADE DE PEDIDOS
with open(PASTA_DADOS / "kpi_quantidade_pedidos.json", "w", encoding="utf-8") as f:
    json.dump({
        "atual": qtd_atual,
        "ano_anterior": qtd_ano_anterior,
        "variacao": round(var_qtd, 1) if var_qtd is not None else None
    }, f, ensure_ascii=False, indent=2)

# TICKET MÉDIO
with open(PASTA_DADOS / "kpi_ticket_medio.json", "w", encoding="utf-8") as f:
    json.dump({
        "atual": round(ticket_atual, 2),
        "ano_anterior": round(ticket_ano_anterior, 2),
        "variacao": round(var_ticket, 1) if var_ticket is not None else None
    }, f, ensure_ascii=False, indent=2)

# ======================================================
# LOG FINAL
# ======================================================
print("✅ KPIs gerados com sucesso")
print(f"💰 Faturamento atual: {fat_atual:,.2f}")
print(f"📦 Quantidade pedidos: {qtd_atual}")
print(f"🎯 Ticket médio: {ticket_atual:,.2f}")
