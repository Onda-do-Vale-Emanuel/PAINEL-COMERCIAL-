import pandas as pd
import json
from datetime import datetime
from pathlib import Path

print("🔄 Iniciando atualização dos KPIs (IGUAL AO EXCEL DEFINITIVO)...")

# ======================================================
# CAMINHOS
# ======================================================
BASE_DIR = Path(__file__).resolve().parents[1]
ARQ_EXCEL = BASE_DIR / "excel" / "PEDIDOS ONDA.xlsx"
PASTA_DADOS = BASE_DIR / "site" / "dados"
PASTA_DADOS.mkdir(exist_ok=True)

# ======================================================
# COLUNAS (base 0)
# ======================================================
COL_TIPO = 3
COL_DATA = 4
COL_VALOR = 7   # 🔥 COLUNA H (use 8 se quiser coluna I)

# ======================================================
# DATA DE REFERÊNCIA
# ======================================================
HOJE = datetime(2026, 1, 26)
ANO_ATUAL = HOJE.year
ANO_ANTERIOR = HOJE.year - 1

# ======================================================
# LEITURA
# ======================================================
df = pd.read_excel(ARQ_EXCEL)
print(f"📄 Linhas lidas: {len(df)}")

# ======================================================
# FILTRO NORMAL
# ======================================================
df = df[df.iloc[:, COL_TIPO].astype(str).str.upper().str.strip() == "NORMAL"]
print(f"✅ NORMAL: {len(df)} linhas")

# ======================================================
# DATA
# ======================================================
df["DATA_OK"] = pd.to_datetime(
    df.iloc[:, COL_DATA],
    errors="coerce",
    dayfirst=True
)
df = df.dropna(subset=["DATA_OK"])

# ======================================================
# VALOR (🔥 SEM CONVERSÃO PARA STRING)
# ======================================================
df["VALOR_OK"] = pd.to_numeric(
    df.iloc[:, COL_VALOR],
    errors="coerce"
).fillna(0)

# ======================================================
# FILTRO DE PERÍODO
# ======================================================
df_atual = df[
    (df["DATA_OK"].dt.year == ANO_ATUAL) &
    (df["DATA_OK"] <= HOJE)
]

df_ano_anterior = df[
    (df["DATA_OK"].dt.year == ANO_ANTERIOR) &
    (df["DATA_OK"] <= HOJE.replace(year=ANO_ANTERIOR))
]

print(f"📆 Pedidos ano atual: {len(df_atual)}")
print(f"📆 Pedidos ano anterior: {len(df_ano_anterior)}")

# ======================================================
# SOMA IGUAL AO EXCEL
# ======================================================
fat_atual = float(df_atual["VALOR_OK"].sum())
fat_ano_anterior = float(df_ano_anterior["VALOR_OK"].sum())

# ======================================================
# QUANTIDADE
# ======================================================
qtd_atual = int(len(df_atual))
qtd_ano_anterior = int(len(df_ano_anterior))

# ======================================================
# VARIAÇÃO
# ======================================================
var_fat = ((fat_atual - fat_ano_anterior) / fat_ano_anterior * 100) if fat_ano_anterior > 0 else None

# ======================================================
# SALVAR JSON
# ======================================================
with open(PASTA_DADOS / "kpi_faturamento.json", "w", encoding="utf-8") as f:
    json.dump({
        "atual": round(fat_atual, 2),
        "ano_anterior": round(fat_ano_anterior, 2),
        "data_atual": HOJE.strftime("%d/%m/%Y"),
        "data_ano_anterior": HOJE.replace(year=ANO_ANTERIOR).strftime("%d/%m/%Y"),
        "variacao": round(var_fat, 1) if var_fat is not None else None
    }, f, ensure_ascii=False, indent=2)

print("✅ KPIs gerados com sucesso (IGUAL AO EXCEL)")
print(f"💰 Atual: {fat_atual:,.2f}")
print(f"💰 Ano anterior: {fat_ano_anterior:,.2f}")
print(f"📦 Pedidos: {qtd_atual}")
