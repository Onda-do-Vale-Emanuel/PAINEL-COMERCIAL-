# ===============================================================
# PAINEL COMERCIAL - INTERFACE FINAL COM SEGURANÇA REAL
# ===============================================================

import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import subprocess
import atualizar_painel_completo as painel
from datetime import datetime
import json
import os
import hashlib
import random
from pathlib import Path

# ===============================================================
# CONFIGURAÇÕES
# ===============================================================

BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
ARQ_LICENCA = BASE_DIR / "licenca_mensal.json"
ARQ_SENHAS_BACKUP = BASE_DIR / "senhas_backup.json"

TENTATIVAS_MAX = 3
tentativas = 0

# ===============================================================
# UTILIDADES DE SEGURANÇA
# ===============================================================

def gerar_hash(texto: str) -> str:
    return hashlib.sha256(texto.encode()).hexdigest()


def gerar_senha_mes() -> str:
    mes = datetime.today().month
    return f"Ondaviperx@{mes:02d}"


# ===============================================================
# GERAR 10 SENHAS BACKUP SEGURAS (RODAR MANUALMENTE SE PRECISAR)
# ===============================================================

def gerar_senhas_backup_seguras():
    """
    Rode essa função MANUALMENTE pelo Python quando quiser gerar novo lote de 10 senhas backup.
    As senhas reais aparecem só no terminal. No JSON ficam apenas os hashes.
    """
    senhas_reais = [str(random.randint(10000000, 99999999)) for _ in range(10)]
    hashes = [gerar_hash(s) for s in senhas_reais]

    with open(ARQ_SENHAS_BACKUP, "w", encoding="utf-8") as f:
        json.dump({"senhas_hash": hashes}, f, indent=2, ensure_ascii=False)

    print("\n=== GUARDE ESTAS SENHAS (NÃO SERÃO MOSTRADAS NOVAMENTE) ===")
    for s in senhas_reais:
        print(s)
    print("===========================================================\n")


# ===============================================================
# LICENÇA MENSAL
# ===============================================================

def licenca_valida() -> bool:
    """Retorna True se já foi liberado neste mês."""
    if not ARQ_LICENCA.exists():
        return False

    try:
        with open(ARQ_LICENCA, "r", encoding="utf-8") as f:
            dados = json.load(f)
    except Exception:
        return False

    mes_atual = datetime.today().strftime("%m/%Y")
    return dados.get("mes") == mes_atual


def salvar_licenca():
    """Marca o mês atual como liberado (para não pedir senha de novo)."""
    mes_atual = datetime.today().strftime("%m/%Y")
    with open(ARQ_LICENCA, "w", encoding="utf-8") as f:
        json.dump({"mes": mes_atual}, f, ensure_ascii=False)


# ===============================================================
# VALIDAR SENHA (MENSAL + BACKUP HASH)
# ===============================================================

def validar_senha(senha_digitada: str) -> bool:
    """Valida senha mensal ou uma senha backup (uso único)."""
    global tentativas

    # 1) Senha mensal normal
    if senha_digitada == gerar_senha_mes():
        salvar_licenca()
        return True

    # 2) Senhas backup com HASH
    if ARQ_SENHAS_BACKUP.exists():
        try:
            with open(ARQ_SENHAS_BACKUP, "r", encoding="utf-8") as f:
                dados = json.load(f)
        except Exception:
            dados = {}

        hashes = dados.get("senhas_hash", [])
        hash_digitado = gerar_hash(senha_digitada)

        if hash_digitado in hashes:
            # uso único → remove
            hashes.remove(hash_digitado)
            with open(ARQ_SENHAS_BACKUP, "w", encoding="utf-8") as f:
                json.dump({"senhas_hash": hashes}, f, indent=2, ensure_ascii=False)
            salvar_licenca()
            return True

    # 3) Se chegou aqui: senha errada
    tentativas += 1

    if tentativas >= TENTATIVAS_MAX:
        messagebox.showerror("Bloqueado", "Faça contato com o Emanuel")
        janela.destroy()
        return False

    messagebox.showerror("Erro", f"Senha incorreta! ({tentativas}/3)")
    return False


# ===============================================================
# MONTAR RESUMO LENDO OS JSONs (GARANTIDO IGUAL AO PAINEL)
# ===============================================================

def montar_resumo_por_json():
    """
    Lê os arquivos JSON gerados por atualizar_painel_completo.py e monta
    o mesmo resumo que você vê no painel: período, pedidos e faturamento.
    Isso independe do que a função main() retorna.
    """
    try:
        # usa o mesmo diretório de dados que o script painel usa
        dados_dir = painel.DADOS_DIR_2

        # Faturamento → tem períodos e valores
        with open(dados_dir / "kpi_faturamento.json", "r", encoding="utf-8") as f:
            kpi_fat = json.load(f)

        # Quantidade de pedidos
        with open(dados_dir / "kpi_quantidade_pedidos.json", "r", encoding="utf-8") as f:
            kpi_ped = json.load(f)

        periodo_atual = f"{kpi_fat.get('inicio_mes', '?')} até {kpi_fat.get('data_atual', '?')}"
        periodo_ant = f"{kpi_fat.get('inicio_mes_anterior', '?')} até {kpi_fat.get('data_ano_anterior', '?')}"

        pedidos_2026 = kpi_ped.get("atual", 0)
        pedidos_2025 = kpi_ped.get("ano_anterior", 0)

        fat_2026 = float(kpi_fat.get("atual", 0) or 0)
        fat_2025 = float(kpi_fat.get("ano_anterior", 0) or 0)

        return {
            "periodo_atual": periodo_atual,
            "periodo_ant": periodo_ant,
            "pedidos_2026": pedidos_2026,
            "pedidos_2025": pedidos_2025,
            "fat_2026": fat_2026,
            "fat_2025": fat_2025,
        }

    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível montar o resumo pelos JSONs:\n{e}")
        return None


# ===============================================================
# EXECUTAR ATUALIZAÇÃO
# ===============================================================

def executar_atualizacao():
    # 🔐 Se ainda não liberou no mês, pede senha
    if not licenca_valida():
        senha_digitada = entry_senha.get().strip()
        if not validar_senha(senha_digitada):
            return
        else:
            # senha ok → esconde bloco de autenticação
            frame_auth.pack_forget()

    usar_auto = var_auto.get()

    try:
        # 1) Roda o cálculo e gera os JSONs
        if usar_auto:
            painel.main()
        else:
            data_ini = entry_ini.get_date().strftime("%d/%m/%Y")
            data_fim = entry_fim.get_date().strftime("%d/%m/%Y")
            painel.main(data_ini, data_fim)

        # 2) Monta o resumo lendo os JSONs
        resumo = montar_resumo_por_json()
        if not resumo:
            return

        mensagem = (
            "ATUALIZAÇÃO CONCLUÍDA!\n\n"
            f"■ Período atual: {resumo['periodo_atual']}\n"
            f"■ Período anterior: {resumo['periodo_ant']}\n\n"
            f"■ Pedidos 2026: {resumo['pedidos_2026']}\n"
            f"■ Pedidos 2025: {resumo['pedidos_2025']}\n\n"
            f"■ Faturamento 2026: R$ {resumo['fat_2026']:,.2f}\n"
            f"■ Faturamento 2025: R$ {resumo['fat_2025']:,.2f}\n"
        )

        messagebox.showinfo("Resumo da Atualização", mensagem)

        # 3) Git automático
        subprocess.run(["git", "add", "."], shell=True)
        subprocess.run(["git", "commit", "-m", "Atualização automática painel"], shell=True)
        subprocess.run(["git", "push"], shell=True)

        messagebox.showinfo("Publicado", "Painel atualizado e enviado ao site!")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


# ===============================================================
# INTERFACE
# ===============================================================

janela = tk.Tk()
janela.title("PAINEL COMERCIAL - Atualização")
janela.geometry("650x420")
janela.resizable(False, False)

titulo = tk.Label(
    janela,
    text="Atualização do Painel Comercial",
    font=("Arial", 18, "bold")
)
titulo.pack(pady=10)

# ---- AUTENTICAÇÃO ----
frame_auth = tk.LabelFrame(janela, text="Autenticação", font=("Arial", 11, "bold"))

if not licenca_valida():
    frame_auth.pack(fill="x", padx=20, pady=10)

tk.Label(
    frame_auth,
    text="Digite a senha mensal:",
    font=("Arial", 11)
).pack(anchor="w", padx=10, pady=5)

entry_senha = tk.Entry(
    frame_auth,
    show="*",
    width=30,
    font=("Arial", 12)
)
entry_senha.pack(padx=10, pady=5)

# ---- PERÍODO ----
frame_periodo = tk.LabelFrame(
    janela,
    text="Período para Atualização",
    font=("Arial", 11, "bold")
)
frame_periodo.pack(fill="x", padx=20, pady=10)

var_auto = tk.BooleanVar(value=True)

tk.Radiobutton(
    frame_periodo,
    text="Período Automático (01 → última data da planilha)",
    variable=var_auto,
    value=True,
    font=("Arial", 11)
).grid(row=0, column=0, sticky="w", padx=10)

tk.Radiobutton(
    frame_periodo,
    text="Período Personalizado",
    variable=var_auto,
    value=False,
    font=("Arial", 11)
).grid(row=1, column=0, sticky="w", padx=10)

tk.Label(frame_periodo, text="Data Inicial:", font=("Arial", 11)).grid(row=2, column=0, sticky="w", padx=10)
entry_ini = DateEntry(frame_periodo, width=12, date_pattern="dd/mm/yyyy")
entry_ini.grid(row=2, column=1, padx=10)

tk.Label(frame_periodo, text="Data Final:", font=("Arial", 11)).grid(row=3, column=0, sticky="w", padx=10)
entry_fim = DateEntry(frame_periodo, width=12, date_pattern="dd/mm/yyyy")
entry_fim.grid(row=3, column=1, padx=10)

# ---- BOTÕES ----
frame_btn = tk.Frame(janela)
frame_btn.pack(pady=20)

btn_atualizar = tk.Button(
    frame_btn,
    text="Iniciar Atualização",
    font=("Arial", 13, "bold"),
    bg="#0a74d4",
    fg="white",
    width=20,
    command=executar_atualizacao
)
btn_atualizar.grid(row=0, column=0, padx=10)

btn_cancelar = tk.Button(
    frame_btn,
    text="Cancelar",
    font=("Arial", 13),
    width=15,
    command=janela.destroy
)
btn_cancelar.grid(row=0, column=1, padx=10)

janela.mainloop()