import pandas as pd
import json
from datetime import datetime
from pathlib import Path
import numpy as np

print("🔄 Iniciando atualização do KPI de quantidade de pedidos...")

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

print(f"📆 Pedidos ano atual: {len(df_atual)}")
print(f"📆 Pedidos ano anterior: {len(df_ano_anterior)}")

# ======================================================
# CÁLCULOS (1 LINHA = 1 PEDIDO)
# ======================================================
qtd_atual = int(len(df_atual))
qtd_ano_anterior = int(len(df_ano_anterior))

if qtd_ano_anterior > 0:
    variacao = float(
        ((qtd_atual - qtd_ano_anterior) / qtd_ano_anterior) * 100
    )
else:
    variacao = None

# ======================================================
# GERAR JSON
# ======================================================
dados = {
    "atual": qtd_atual,
    "ano_anterior": qtd_ano_anterior,
    "data_atual": HOJE.strftime("%d/%m/%Y"),
    "data_ano_anterior": HOJE.replace(year=ANO_ANTERIOR).strftime("%d/%m/%Y"),
    "variacao": round(variacao, 1) if variacao is not None else None
}

with open(PASTA_DADOS / "kpi_quantidade_pedidos.json", "w", encoding="utf-8") as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

# ======================================================
# LOG FINAL
# ======================================================
print("✅ KPI quantidade de pedidos gerado com sucesso")
print(f"📦 Pedidos até {dados['data_atual']}: {qtd_atual}")
print(f"📦 Pedidos até {dados['data_ano_anterior']}: {qtd_ano_anterior}")
print(f"📈 Variação: {dados['variacao'] if dados['variacao'] is not None else '--'}%")
