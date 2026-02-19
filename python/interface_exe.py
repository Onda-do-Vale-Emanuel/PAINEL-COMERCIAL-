# ===============================================================
#  PAINEL COMERCIAL - INTERFACE FINAL (TKCALENDAR)
#  Compatível com atualizar_painel_completo.py revisado
# ===============================================================

import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import subprocess
import atualizar_painel_completo as painel
from datetime import datetime

# ===============================================================
# GERA SENHA DO MÊS
# ===============================================================
def gerar_senha_mes():
    mes = datetime.today().month
    return f"Ondaviperx@{mes:02d}"


# ===============================================================
# EXECUTA O PAINEL + PUSH PARA GITHUB
# ===============================================================
def executar_atualizacao():

    senha_digitada = entry_senha.get().strip()
    senha_correta = gerar_senha_mes()

    if senha_digitada != senha_correta:
        messagebox.showerror("Erro", "Senha incorreta!")
        return

    try:
        # =======================================================
        # PERÍODO AUTOMÁTICO OU PERSONALIZADO
        # =======================================================
        if var_auto.get():
            resumo = painel.main()
        else:
            data_ini = entry_ini.get_date().strftime("%d/%m/%Y")
            data_fim = entry_fim.get_date().strftime("%d/%m/%Y")
            resumo = painel.main(data_ini, data_fim)

        if resumo is None:
            raise Exception("Função main() retornou vazio.")

        # =======================================================
        # RESUMO PARA O USUÁRIO
        # =======================================================
        atual = resumo["atual"]
        anterior = resumo["anterior"]

        mensagem = (
            f"ATUALIZAÇÃO CONCLUÍDA!\n\n"
            f"📅 Período atual: {atual['inicio']} até {atual['fim']}\n"
            f"📅 Ano anterior:  {anterior['inicio']} até {anterior['fim']}\n\n"
            f"📌 Pedidos 2026: {atual['pedidos']}\n"
            f"📌 Pedidos 2025: {anterior['pedidos']}\n\n"
            f"💰 Faturamento 2026: R$ {atual['fat']:,}\n"
            f"💰 Faturamento 2025: R$ {anterior['fat']:,}\n"
        )

        messagebox.showinfo("Resumo da Atualização", mensagem)

        # =======================================================
        # PUSH AUTOMÁTICO PARA GITHUB
        # =======================================================
        if resumo["push"]:
            messagebox.showinfo("Sucesso", "Dados enviados ao site com sucesso!")
        else:
            messagebox.showwarning(
                "Aviso",
                "Painel atualizado, mas não foi possível enviar para o site.\n"
                "Verifique se o Git está configurado."
            )

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


# ===============================================================
# INTERFACE TKINTER
# ===============================================================

janela = tk.Tk()
janela.title("PAINEL COMERCIAL - Atualização")
janela.geometry("650x430")
janela.resizable(False, False)

titulo = tk.Label(janela, text="Atualização do Painel Comercial",
                  font=("Arial", 18, "bold"))
titulo.pack(pady=10)

# ----------------------------- AUTENTICAÇÃO ---------------------
frame_auth = tk.LabelFrame(janela, text="Autenticação", font=("Arial", 11, "bold"))
frame_auth.pack(fill="x", padx=20, pady=10)

tk.Label(frame_auth, text="Digite a senha mensal:", font=("Arial", 11)).pack(anchor="w", padx=10, pady=5)

entry_senha = tk.Entry(frame_auth, show="*", width=30, font=("Arial", 12))
entry_senha.pack(padx=10, pady=5)

# ----------------------------- PERÍODO ---------------------------
frame_periodo = tk.LabelFrame(janela, text="Período para Atualização", font=("Arial", 11, "bold"))
frame_periodo.pack(fill="x", padx=20, pady=10)

var_auto = tk.BooleanVar(value=True)

tk.Radiobutton(frame_periodo, text="Período Automático (01 → última data da planilha)",
               variable=var_auto, value=True,
               font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=10)

tk.Radiobutton(frame_periodo, text="Período Personalizado",
               variable=var_auto, value=False,
               font=("Arial", 11)).grid(row=1, column=0, sticky="w", padx=10)

tk.Label(frame_periodo, text="Data Inicial:", font=("Arial", 11)).grid(row=2, column=0, sticky="w", padx=10)
entry_ini = DateEntry(frame_periodo, width=12, date_pattern="dd/mm/yyyy")
entry_ini.grid(row=2, column=1, padx=10)

tk.Label(frame_periodo, text="Data Final:", font=("Arial", 11)).grid(row=3, column=0, sticky="w", padx=10)
entry_fim = DateEntry(frame_periodo, width=12, date_pattern="dd/mm/yyyy")
entry_fim.grid(row=3, column=1, padx=10)

# ----------------------------- BOTÕES ---------------------------
frame_btn = tk.Frame(janela)
frame_btn.pack(pady=25)

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
