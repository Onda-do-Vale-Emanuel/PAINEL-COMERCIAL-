import pandas as pd
import json
import re
from datetime import datetime

CAMINHO_EXCEL = "excel/PEDIDOS ONDA.xlsx"

# ======================================================
# CONVERTE DATA/HORA DO EXCEL (ex: 29/10/1900 01:12:00) EM NÚMERO (ex: 303,05)
# ======================================================
def excel_datetime_para_numero(dt):
    # Excel (Windows) usa como base 1899-12-30 para serial date
    base = datetime(1899, 12, 30)
    return (dt - base).total_seconds() / 86400.0

# ======================================================
# FUNÇÃO DEFINITIVA PARA LER NÚMEROS BRASILEIROS + TRATAR DATETIME EM COLUNAS NUMÉRICAS
# ======================================================
def limpar_numero(v):
    if pd.isna(v):
        return 0.0

    # Se já for número
    if isinstance(v, (int, float)):
        return float(v)

    # Se for datetime vindo do Excel (ex.: 29/10/1900 01:12:00)
    if isinstance(v, (datetime, pd.Timestamp)):
        dt = pd.to_datetime(v).to_pydatetime()
        return float(excel_datetime_para_numero(dt))

    s = str(v).strip()

    # Tentar detectar string datetime e converter para serial (evita casos tipo "1900-10-29011200")
    try:
        # se parece data/hora, converte
        dt = pd.to_datetime(s, errors="raise", dayfirst=True)
        # se tinha traço/barra/tempo no texto, tratamos como data
        if any(ch in s for ch in ["/", "-", ":"]):
            return float(excel_datetime_para_numero(dt.to_pydatetime()))
    except Exception:
        pass

    # Limpa tudo que não for número/sinal/ponto/vírgula
    s = re.sub(r"[^0-9,.\-]", "", s)

    if s in ["", "-", ",", ".", ",-", ".-"]:
        return 0.0

    # Caso "1.234.567,89"
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    # Caso "123,45"
    if "," in s and "." not in s:
        try:
            return float(s.replace(",", "."))
        except Exception:
            return 0.0

    # Caso "12.345" (milhar) ou "123.4"
    if "." in s and "," not in s:
        partes = s.split(".")
        # se termina com 3 dígitos, provavelmente é milhar
        if len(partes[-1]) == 3:
            s = s.replace(".", "")
        try:
            return float(s)
        except Exception:
            return 0.0

    try:
        return float(s)
    except Exception:
        return 0.0

# ======================================================
# CARREGA PLANILHA + NORMALIZA + FILTRA APENAS PEDIDOS VÁLIDOS
# ======================================================
def carregar():
    df = pd.read_excel(CAMINHO_EXCEL)
    df.columns = df.columns.str.upper().str.strip()

    # Obrigatórias
    obrig = ["DATA", "PEDIDO", "TIPO DE PEDIDO", "VALOR COM IPI", "KG", "TOTAL M2"]
    for c in obrig:
        if c not in df.columns:
            raise Exception(f"❌ Coluna obrigatória não encontrada: {c}")

    # DATA
    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")
    df = df[df["DATA"].notna()]

    # Números
    for col in ["VALOR COM IPI", "KG", "TOTAL M2"]:
        df[col] = df[col].apply(limpar_numero)

    # Pedido numérico (aceita 38.504 e 38.504,00)
    df["PEDIDO_NUM"] = df["PEDIDO"].apply(limpar_numero)

    # Filtrar TIPO DE PEDIDO = Normal
    df["TIPO DE PEDIDO"] = df["TIPO DE PEDIDO"].astype(str).str.strip().str.upper()
    df = df[df["TIPO DE PEDIDO"] == "NORMAL"]

    # Filtrar pedidos válidos (faixa que você definiu)
    df = df[(df["PEDIDO_NUM"] >= 30000) & (df["PEDIDO_NUM"] <= 50000)]

    # Remover duplicidade por segurança (você disse que não repete)
    df = df.drop_duplicates(subset=["PEDIDO_NUM"])

    return df

# ======================================================
# DEFINE PERÍODOS (CÁLCULO REAL) E (EXIBIÇÃO NO SITE)
# Regras:
# - CÁLCULO ATUAL: primeira data real com pedido no mês → última data real com pedido no mês
# - CÁLCULO ANTERIOR: mesma regra, mas limitado até o DIA do fim do atual (ex: dia 05)
# - EXIBIÇÃO: sempre "01/mm/aaaa até dd/mm/aaaa" (dia final do mês atual)
# ======================================================
def obter_periodos(df):
    ultima_real = df["DATA"].max()
    ano = ultima_real.year
    mes = ultima_real.month
    dia_final = ultima_real.day

    # mês atual (após filtros)
    df_mes_atual = df[(df["DATA"].dt.year == ano) & (df["DATA"].dt.month == mes)]
    inicio_real_atual = df_mes_atual["DATA"].min()
    fim_real_atual = df_mes_atual["DATA"].max()

    # ano anterior (mesmo mês, até o mesmo dia do mês do atual)
    ano_ant = ano - 1
    alvo_fim_ant = datetime(ano_ant, mes, dia_final)

    df_mes_ant = df[(df["DATA"].dt.year == ano_ant) & (df["DATA"].dt.month == mes)]
    if len(df_mes_ant) == 0:
        # se não existir nada no mês do ano anterior
        inicio_real_ant = datetime(ano_ant, mes, 1)
        fim_real_ant = datetime(ano_ant, mes, 1)
    else:
        # limita até o dia_final, mas se não houver no dia, pega o último <= dia_final
        df_ant_limitado = df_mes_ant[df_mes_ant["DATA"] <= alvo_fim_ant]
        if len(df_ant_limitado) > 0:
            fim_real_ant = df_ant_limitado["DATA"].max()
        else:
            fim_real_ant = df_mes_ant["DATA"].max()

        inicio_real_ant = df_mes_ant["DATA"].min()

    # EXIBIÇÃO NO SITE: sempre 01/mm até dia_final (mesmo se não teve pedido no dia 01)
    inicio_exib_atual = datetime(ano, mes, 1)
    fim_exib_atual = datetime(ano, mes, dia_final)

    inicio_exib_ant = datetime(ano_ant, mes, 1)
    fim_exib_ant = datetime(ano_ant, mes, dia_final)

    return {
        "atual": {
            "inicio_real": inicio_real_atual,
            "fim_real": fim_real_atual,
            "inicio_exib": inicio_exib_atual,
            "fim_exib": fim_exib_atual,
        },
        "anterior": {
            "inicio_real": inicio_real_ant,
            "fim_real": fim_real_ant,
            "inicio_exib": inicio_exib_ant,
            "fim_exib": fim_exib_ant,
        },
        "ultima_data": ultima_real
    }

# ======================================================
# RESUMO DO PERÍODO (sempre usa período REAL)
# ======================================================
def resumo(df, inicio, fim):
    d = df[(df["DATA"] >= inicio) & (df["DATA"] <= fim)]

    pedidos = len(d)
    fat = float(d["VALOR COM IPI"].sum())
    kg = float(d["KG"].sum())
    m2 = float(d["TOTAL M2"].sum())

    ticket = fat / pedidos if pedidos else 0.0
    preco_kg = fat / kg if kg else 0.0
    preco_m2 = fat / m2 if m2 else 0.0

    return {
        "pedidos": pedidos,
        "fat": fat,
        "kg": kg,
        "m2": m2,
        "ticket": ticket,
        "preco_kg": preco_kg,
        "preco_m2": preco_m2
    }

# ======================================================
# SALVAR JSON
# ======================================================
def salvar(nome, dados):
    with open(f"dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    with open(f"site/dados/{nome}", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

# ======================================================
# EXECUÇÃO
# ======================================================
if __name__ == "__main__":
    df = carregar()
    p = obter_periodos(df)

    atual = resumo(df, p["atual"]["inicio_real"], p["atual"]["fim_real"])
    anterior = resumo(df, p["anterior"]["inicio_real"], p["anterior"]["fim_real"])

    # FATURAMENTO
    salvar("kpi_faturamento.json", {
        "atual": round(atual["fat"], 2),
        "ano_anterior": round(anterior["fat"], 2),
        "variacao": ((atual["fat"] / anterior["fat"]) - 1) * 100 if anterior["fat"] else 0,
        "inicio_mes": p["atual"]["inicio_exib"].strftime("%d/%m/%Y"),
        "data_atual": p["atual"]["fim_exib"].strftime("%d/%m/%Y"),
        "inicio_mes_anterior": p["anterior"]["inicio_exib"].strftime("%d/%m/%Y"),
        "data_ano_anterior": p["anterior"]["fim_exib"].strftime("%d/%m/%Y")
    })

    # QUANTIDADE PEDIDOS
    salvar("kpi_quantidade_pedidos.json", {
        "atual": atual["pedidos"],
        "ano_anterior": anterior["pedidos"],
        "variacao": ((atual["pedidos"] / anterior["pedidos"]) - 1) * 100 if anterior["pedidos"] else 0
    })

    # KG TOTAL
    salvar("kpi_kg_total.json", {
        "atual": round(atual["kg"], 2),
        "ano_anterior": round(anterior["kg"], 2),
        "variacao": ((atual["kg"] / anterior["kg"]) - 1) * 100 if anterior["kg"] else 0
    })

    # TICKET MÉDIO
    salvar("kpi_ticket_medio.json", {
        "atual": round(atual["ticket"], 2),
        "ano_anterior": round(anterior["ticket"], 2),
        "variacao": ((atual["ticket"] / anterior["ticket"]) - 1) * 100 if anterior["ticket"] else 0
    })

    # PREÇO MÉDIO (AGORA COM COMPARATIVO 2026 vs 2025)
    salvar("kpi_preco_medio.json", {
        "atual": {
            "preco_medio_kg": round(atual["preco_kg"], 2),
            "preco_medio_m2": round(atual["preco_m2"], 2),
            "total_kg": round(atual["kg"], 2),
            "total_m2": round(atual["m2"], 2),
            "data": p["atual"]["fim_exib"].strftime("%d/%m/%Y")
        },
        "ano_anterior": {
            "preco_medio_kg": round(anterior["preco_kg"], 2),
            "preco_medio_m2": round(anterior["preco_m2"], 2),
            "total_kg": round(anterior["kg"], 2),
            "total_m2": round(anterior["m2"], 2),
            "data": p["anterior"]["fim_exib"].strftime("%d/%m/%Y")
        }
    })

    print("=====================================")
    print(" ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!")
    print("=====================================")
    print(f"📅 Última data encontrada: {p['ultima_data']}")
    print(f"Pedidos atuais: {atual['pedidos']}")
    print(f"Pedidos ano anterior: {anterior['pedidos']}")
