import pandas as pd

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

df = pd.read_excel(CAMINHO_EXCEL)
df.columns = df.columns.str.upper().str.strip()

print("\n================ COLUNAS ENCONTRADAS ================")
print(df.columns.tolist())

# Converter DATA
df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
df = df[df["DATA"].notna()]

print("\n================ RESUMO GERAL ========================")
print("Total de linhas na planilha:", len(df))

# Filtrar apenas fevereiro de 2026
df_2026 = df[
    (df["DATA"].dt.year == 2026) &
    (df["DATA"].dt.month == 2)
]

print("\n================ DADOS 2026 FEVEREIRO =================")
print("Total de linhas 2026 Fevereiro:", len(df_2026))

if len(df_2026) > 0:
    print("Primeira data encontrada:", df_2026["DATA"].min())
    print("Última data encontrada:", df_2026["DATA"].max())

    print("\nContagem por dia:")
    print(df_2026["DATA"].dt.day.value_counts().sort_index())

    if "TIPO DE PEDIDO" in df.columns:
        print("\nTipos de pedido encontrados:")
        print(df_2026["TIPO DE PEDIDO"].value_counts())

    if "PEDIDO" in df.columns:
        print("\nExemplo dos últimos 10 pedidos:")
        print(df_2026[["PEDIDO", "DATA"]].tail(10))

print("\n======================================================")
