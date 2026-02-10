import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ======================================================
# 🔥 FUNÇÃO DEFINITIVA PARA LER NÚMEROS BRASILEIROS
# ======================================================
def limpar_numero(valor):
    """Converte qualquer entrada para número adequado.
       Trata:
       - números BR
       - números com ponto
       - datas convertidas para número Excel (ex: 29/10/1900 → 303,05)
       - valores ISO de data
       - texto contendo moedas, símbolos etc."""
    
    if pd.isna(valor):
        return 0.0

    # 1) SE FOR DATA REAL (Timestamp)
    if isinstance(valor, pd.Timestamp) or isinstance(valor, datetime):
        base = datetime(1899, 12, 30)
        delta = valor - base
        return round(delta.days + delta.seconds / 86400, 2)

    v = str(valor).strip()

    # 2) DETECTA TEXTO EM FORMATO DE DATA
    padrao_data_texto = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})(.*\d{1,2}:\d{2})?"
    padrao_data_iso = r"^\d{4}-\d{2}-\d{2}"

    if re.search(padrao_data_texto, v) or re.match(padrao_data_iso, v):
        try:
            dt = pd.to_datetime(v, errors="coerce")
            if pd.notna(dt):
                base = datetime(1899, 12, 30)
                delta = dt - base
                return round(delta.days + delta.seconds / 86400, 2)
        except:
            pass

    # 3) LIMPA CARACTERES
    v = re.sub(r"[^0-9,.-]", "", v)

    if v in ["", "-", ".", ","]:
        return 0.0

    # 4) FORMATO BRISIEN (1.234,56)
    if "." in v and "," in v:
        v = v.replace(".", "").replace(",", ".")
        return float(v)

    # 5) FORMATO BR (123,45)
    if "," in v:
        return float(v.replace(",", "."))

    # 6) FORMATO US (123.45)
    try:
        return float(v)
    except:
        return 0.0


# ======================================================
# 🔥 CARREGAR PLANILHA E LIMPAR DADOS
# ======================================================
def carregar():
    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.upper().str.strip()

    # LIMPA COLUNAS NUMÉRICAS
    colunas_numericas = ["VALOR TOTAL", "VALOR PRODUTO", "VALOR EMBALAGEM",
                         "VALOR COM IPI", "KG", "TOTAL M2"]

    for col in colunas_numericas:
        if col in df.columns:
            df[col] = df[col].apply(limpar_numero).round(2)

    # CONVERTE DATA
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    # CONVERTE NÚMERO DO PEDIDO (COM DATAS E FORMATOS BIZARROS)
    df["PEDIDO_NUM"] = df["PEDIDO"].apply(limpar_numero)

    # REGRA: PEDIDOS ENTRE 30 MIL E 50 MIL
    df = df[(df["PEDIDO_NUM"] >= 30000) & (df["PEDIDO_NUM"] <= 50000)]

    return df


# ======================================================
# 🔥 PERÍODOS (MESMO PADRÃO PARA ATUAL E ANO ANTERIOR)
# ======================================================
def obter_periodos(df):
    ultima_data = df["DATA"].max()

    inicio_atual = ultima_data.replace(day=1)

    ano_ant = ultima_data.year - 1
    inicio_ant = inicio_atual.replace(year=ano_ant)

    df_ant_mes = df[(df["DATA"].dt.year == ano_ant) &
                    (df["DATA"].dt.month == ultima_data.month)]

    fim_ant = df_ant_mes["DATA"].max() if len(df_ant_mes) else inicio_ant

    return (inicio_atual, ultima_data), (inicio_ant, fim_ant)


# ======================================================
# 🔥 RESUMO GENÉRICO
# ======================================================
def resumo(df, inicio, fim):
    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)]

    total_valor = d["VALOR COM IPI"].sum()
    total_kg = d["KG"].sum()
    total_m2 = d["TOTAL M2"].sum()
    pedidos = len(d)

    ticket = total_valor / pedidos if pedidos else 0
    preco_kg = total_valor / total_kg if total_kg else 0
    preco_m2 = total_valor / total_m2 if total_m2 else 0

    return {
        "pedidos": pedidos,
        "fat": total_valor,
        "kg": round(total_kg, 2),
        "m2": round(total_m2, 2),
        "ticket": round(ticket, 2),
        "preco_kg": round(preco_kg, 2),
        "preco_m2": round(preco_m2, 2),
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y")
    }


# ======================================================
# 🔥 SALVAR JSON (PASTA raiz + site/)
# ======================================================
def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


# ======================================================
# 🔥 PRINCIPAL
# ======================================================
if __name__ == "__main__":
    df = carregar()
    (inicio_atual, fim_atual), (inicio_ant, fim_ant) = obter_periodos(df)

    atual = resumo(df, inicio_atual, fim_atual)
    anterior = resumo(df, inicio_ant, fim_ant)

    # FATURAMENTO
    salvar("kpi_faturamento.json", {
        "atual": atual["fat"],
        "ano_anterior": anterior["fat"],
        "variacao": ((atual["fat"] / anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": atual["inicio"],
        "data_atual": atual["fim"],
        "inicio_mes_anterior": anterior["inicio"],
        "data_ano_anterior": anterior["fim"]
    })

    # QUANTIDADE
    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"] / anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    })

    # KG
    salvar("kpi_kg_total.json", {
        "atual": atual["kg"],
        "ano_anterior": anterior["kg"],
        "variacao": ((atual["kg"] / anterior["kg"]) - 1) * 100 if anterior["kg"] else 0
    })

    # TICKET
    salvar("kpi_ticket_medio.json", {
        "atual": atual["ticket"],
        "ano_anterior": anterior["ticket"],
        "variacao": ((atual["ticket"] / anterior["ticket"]) - 1) * 100 if anterior["ticket"] else 0
    })

    # PREÇO MÉDIO (ATUAL E ANTERIOR)
    salvar("kpi_preco_medio.json", {
        "atual": {
            "preco_medio_kg": atual["preco_kg"],
            "preco_medio_m2": atual["preco_m2"],
            "data": atual["fim"]
        },
        "ano_anterior": {
            "preco_medio_kg": anterior["preco_kg"],
            "preco_medio_m2": anterior["preco_m2"],
            "data": anterior["fim"]
        }
    })

    print("\n=====================================")
    print(" ATUALIZAÇÃO CONCLUÍDA COM SUCESSO! ")
    print("=====================================\n")
