import pandas as pd
import json
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

JSON_LOCAL = "dados/kpi_preco_medio.json"
JSON_SITE = "site/dados/kpi_preco_medio.json"

df = pd.read_excel(CAMINHO_EXCEL, engine="openpyxl")

df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

agora = datetime.now()
mes = agora.month
ano = agora.year

df_mes = df[(df["Data"].dt.month == mes) & (df["Data"].dt.year == ano)]

total_valor_com_ipi = df_mes["Valor Com IPI"].sum()
total_kg = df_mes["Kg"].sum()
total_m2 = df_mes["Total m2"].sum()

media_kg = round(total_valor_com_ipi / total_kg, 2) if total_kg > 0 else 0
media_m2 = round(total_valor_com_ipi / total_m2, 2) if total_m2 > 0 else 0

resultado = {
    "preco_medio_kg": media_kg,
    "preco_medio_m2": media_m2,
    "total_kg": round(total_kg, 2),
    "total_m2": round(total_m2, 2),
    "data": agora.strftime("%d/%m/%Y")
}

for arquivo in [JSON_LOCAL, JSON_SITE]:
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(resultado, f, indent=2, ensure_ascii=False)

print("Preço médio gerado:")
print(json.dumps(resultado, indent=2, ensure_ascii=False))
