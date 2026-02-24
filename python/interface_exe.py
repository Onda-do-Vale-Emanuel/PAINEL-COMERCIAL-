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
# SEGURANÇA
# ===============================================================

def gerar_hash(texto):
    return hashlib.sha256(texto.encode()).hexdigest()

def gerar_senha_mes():
    mes = datetime.today().month
    return f"Ondaviperx@{mes:02d}"

# ===============================================================
# LICENÇA
# ===============================================================

def licenca_valida():
    if not ARQ_LICENCA.exists():
        return False

    try:
        with open(ARQ_LICENCA, "r") as f:
            dados = json.load(f)
    except:
        return False

    mes_atual = datetime.today().strftime("%m/%Y")
    return dados.get("mes") == mes_atual


def salvar_licenca():
    mes_atual = datetime.today().strftime("%m/%Y")
    with open(ARQ_LICENCA, "w") as f:
        json.dump({"mes": mes_atual}, f)

# ===============================================================
# VALIDAR SENHA
# ===============================================================

def validar_senha(senha_digitada):
    global tentativas

    if senha_digitada == gerar_senha_mes():
        salvar_licenca()
        return True

    if ARQ_SENHAS_BACKUP.exists():
        try:
            with open(ARQ_SENHAS_BACKUP, "r") as f:
                dados = json.load(f)
        except:
            dados = {}

        hashes = dados.get("senhas_hash", [])
        if gerar_hash(senha_digitada) in hashes:
            hashes.remove(gerar_hash(senha_digitada))
            with open(ARQ_SENHAS_BACKUP, "w") as f:
                json.dump({"senhas_hash": hashes}, f, indent=2)
            salvar_licenca()
            return True

    tentativas += 1

    if tentativas >= TENTATIVAS_MAX:
        messagebox.showerror("Bloqueado", "Faça contato com o Emanuel")
        janela.destroy()
        return False

    messagebox.showerror("Erro", f"Senha incorreta! ({tentativas}/3)")
    return False

# ===============================================================
# EXECUTAR ATUALIZAÇÃO
# ===============================================================

def executar_atualizacao():

    if not licenca_valida():
        senha_digitada = entry_senha.get().strip()
        if not validar_senha(senha_digitada):
            return
        else:
            frame_auth.pack_forget()

    try:

        if var_auto.get():
            resumo = painel.main()
        else:
            data_ini = entry_ini.get_date().strftime("%d/%m/%Y")
            data_fim = entry_fim.get_date().strftime("%d/%m/%Y")
            resumo = painel.main(data_ini, data_fim)

        # 🔒 Garantia contra erro de chave
        periodo_atual = resumo.get("periodo_atual", "Não informado")
        periodo_ant = resumo.get("periodo_ant", "Não informado")
        pedidos_2026 = resumo.get("pedidos_2026", 0)
        pedidos_2025 = resumo.get("pedidos_2025", 0)
        fat_2026 = float(resumo.get("fat_2026", 0))
        fat_2025 = float(resumo.get("fat_2025", 0))

        mensagem = (
            f"ATUALIZAÇÃO CONCLUÍDA!\n\n"
            f"■ Período atual: {periodo_atual}\n"
            f"■ Período anterior: {periodo_ant}\n\n"
            f"■ Pedidos 2026: {pedidos_2026}\n"
            f"■ Pedidos 2025: {pedidos_2025}\n\n"
            f"■ Faturamento 2026: R$ {fat_2026:,.2f}\n"
            f"■ Faturamento 2025: R$ {fat_2025:,.2f}\n"
        )

        messagebox.showinfo("Resumo da Atualização", mensagem)

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

frame_auth = tk.LabelFrame(janela, text="Autenticação", font=("Arial", 11, "bold"))

if not licenca_valida():
    frame_auth.pack(fill="x", padx=20, pady=10)

tk.Label(frame_auth, text="Digite a senha mensal:", font=("Arial", 11)).pack(anchor="w", padx=10, pady=5)

entry_senha = tk.Entry(frame_auth, show="*", width=30, font=("Arial", 12))
entry_senha.pack(padx=10, pady=5)

frame_periodo = tk.LabelFrame(janela, text="Período para Atualização", font=("Arial", 11, "bold"))
frame_periodo.pack(fill="x", padx=20, pady=10)

var_auto = tk.BooleanVar(value=True)

tk.Radiobutton(frame_periodo,
               text="Período Automático (01 → última data da planilha)",
               variable=var_auto,
               value=True,
               font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=10)

tk.Radiobutton(frame_periodo,
               text="Período Personalizado",
               variable=var_auto,
               value=False,
               font=("Arial", 11)).grid(row=1, column=0, sticky="w", padx=10)

tk.Label(frame_periodo, text="Data Inicial:", font=("Arial", 11)).grid(row=2, column=0, sticky="w", padx=10)
entry_ini = DateEntry(frame_periodo, width=12, date_pattern="dd/mm/yyyy")
entry_ini.grid(row=2, column=1, padx=10)

tk.Label(frame_periodo, text="Data Final:", font=("Arial", 11)).grid(row=3, column=0, sticky="w", padx=10)
entry_fim = DateEntry(frame_periodo, width=12, date_pattern="dd/mm/yyyy")
entry_fim.grid(row=3, column=1, padx=10)

frame_btn = tk.Frame(janela)
frame_btn.pack(pady=20)

btn_atualizar = tk.Button(frame_btn,
                          text="Iniciar Atualização",
                          font=("Arial", 13, "bold"),
                          bg="#0a74d4",
                          fg="white",
                          width=20,
                          command=executar_atualizacao)
btn_atualizar.grid(row=0, column=0, padx=10)

btn_cancelar = tk.Button(frame_btn,
                         text="Cancelar",
                         font=("Arial", 13),
                         width=15,
                         command=janela.destroy)
btn_cancelar.grid(row=0, column=1, padx=10)

janela.mainloop()