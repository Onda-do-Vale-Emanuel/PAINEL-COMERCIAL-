import pandas as pd
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

def limpar_numero(v):
    if pd.isna(v):
        return 0
    v = str(v).strip()
    v = re.sub(r"[^0-9,.-]", "", v)
    if v == "":
        return 0
    if "." in v and "," in v:
        try:
            return float(v.replace(".", "").replace(",", "."))
        except:
            return 0
    if "," in v:
        try:
            return float(v.replace(",", "."))
        except:
            return 0
    try:
        return float(v)
    except:
        return 0

def carregar():
    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.upper().str.strip()
    for col in ["VALOR COM IPI", "KG", "TOTAL M2"]:
        if col in df.columns:
            df[col] = df[col].apply(limpar_numero)
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]
    df["PEDIDO_NUM"] = df["PEDIDO"].apply(lambda x: limpar_numero(x))
    df = df[(df["PEDIDO_NUM"] >= 30000) & (df["PEDIDO_NUM"] <= 50000)]
    return df

def resumo(df, inicio, fim):
    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)]
    return {
        "pedidos": len(d),
        "fat": d["VALOR COM IPI"].sum(),
        "kg": d["KG"].sum(),
        "m2": d["TOTAL M2"].sum(),
    }

if __name__ == "__main__":
    df = carregar()
    ultima = df["DATA"].max()
    inicio_atual = ultima.replace(day=1)

    ano_ant = ultima.year - 1
    inicio_ant = inicio_atual.replace(year=ano_ant)

    df_ant_mes = df[
        (df["DATA"].dt.year == ano_ant) &
        (df["DATA"].dt.month == ultima.month)
    ]
    if len(df_ant_mes) > 0:
        fim_ant = df_ant_mes["DATA"].max()
    else:
        fim_ant = inicio_ant

    res_atual = resumo(df, inicio_atual, ultima)
    res_ant = resumo(df, inicio_ant, fim_ant)

    print("\n==================== RESULTADOS ====================")
    print(f"📅 ATUAL    : {inicio_atual.strftime('%d/%m/%Y')} → {ultima.strftime('%d/%m/%Y')}")
    print(res_atual)
    print()
    print(f"📅 ANTERIOR : {inicio_ant.strftime('%d/%m/%Y')} → {fim_ant.strftime('%d/%m/%Y')}")
    print(res_ant)
    print("====================================================\n")
