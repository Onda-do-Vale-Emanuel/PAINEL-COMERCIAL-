import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"
CAMINHO_JSON = "dados/kpi_preco_medio.json"
CAMINHO_SITE = "site/dados/kpi_preco_medio.json"

# ======================================================
# 🔥 FUNÇÃO DEFINITIVA PARA LER NÚMEROS BRASILEIROS
# ======================================================
def limpar_numero(valor):
    if pd.isna(valor):
        return 0.0

    v = str(valor).strip()
    v = re.sub(r"[^0-9,.-]", "", v)

    if v in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0

    # Caso tenha "." e "," → formato 1.234.567,89
    if "." in v and "," in v:
        v = v.replace(".", "")
        v = v.replace(",", ".")
        try:
            return float(v)
        except:
            return 0.0

    # apenas vírgula → decimal
    if "," in v:
        try:
            return float(v.replace(",", "."))
        except:
            return 0.0

    # apenas ponto
    if "." in v:
        partes = v.split(".")
        # se terminar com 3 dígitos → milhar
        if len(partes[-1]) == 3:
            try:
                return float(v.replace(".", ""))
            except:
                return 0.0
        # senão → decimal real
        try:
            return float(v)
        except:
            return 0.0

    # apenas números puros
    try:
        return float(v)
    except:
        return 0.0


# ======================================================
# ✔ CARREGAR EXCEL COM LIMPEZA REAL
# ======================================================
def carregar_excel():
    df = pd.read_excel(CAMINHO_EXCEL)

    df.columns = df.columns.str.upper().str.strip()

    obrig = ["DATA", "VALOR COM IPI", "KG", "TOTAL M2"]
    for c in obrig:
        if c not in df.columns:
            raise Exception(f"❌ Coluna obrigatória ausente: {c}")

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")

    df["VALOR COM IPI"] = df["VALOR COM IPI"].apply(limpar_numero)
    df["KG"] = df["KG"].apply(limpar_numero)
    df["TOTAL M2"] = df["TOTAL M2"].apply(limpar_numero)

    return df


# ======================================================
# ✔ ACHAR A ÚLTIMA DATA REAL DO EXCEL
# ======================================================
def obter_data_ref(df):
    datas = df["DATA"].dropna()
    if datas.empty:
        raise Exception("❌ Nenhuma data válida encontrada.")
    return datas.max()


# ======================================================
# ✔ CALCULAR PREÇO MÉDIO CORRETAMENTE
# ======================================================
def calcular_preco_medio(df, data_ref):
    mes = data_ref.month
    ano = data_ref.year

    df_mes = df[(df["DATA"].dt.month == mes) & (df["DATA"].dt.year == ano)]

    total_valor = df_mes["VALOR COM IPI"].sum()
    total_kg = df_mes["KG"].sum()
    total_m2 = df_mes["TOTAL M2"].sum()

    preco_kg = round(total_valor / total_kg, 2) if total_kg else 0
    preco_m2 = round(total_valor / total_m2, 2) if total_m2 else 0

    return {
        "preco_medio_kg": preco_kg,
        "preco_medio_m2": preco_m2,
        "total_kg": round(total_kg, 2),
        "total_m2": round(total_m2, 2),
        "data": data_ref.strftime("%d/%m/%Y")
    }


# ======================================================
# ✔ SALVAR JSON
# ======================================================
def salvar(dados):
    with open(CAMINHO_JSON, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    with open(CAMINHO_SITE, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


# ======================================================
# ✔ EXECUTAR
# ======================================================
if __name__ == "__main__":
    df = carregar_excel()
    data_ref = obter_data_ref(df)

    print("📅 Última data encontrada:", data_ref)

    preco = calcular_preco_medio(df, data_ref)

    print(json.dumps(preco, ensure_ascii=False, indent=2))

    salvar(preco)

    print("✓ JSON gerado com sucesso.")
