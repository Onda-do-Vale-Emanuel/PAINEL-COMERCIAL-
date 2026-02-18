import tkinter as tk
from tkinter import messagebox
import subprocess
import sys
import os

# Caminho do script principal (motor)
SCRIPT_PRINCIPAL = os.path.join(os.path.dirname(__file__), "atualizar_painel_completo.py")


def executar_atualizacao():
    data_inicio = entry_inicio.get().strip()
    data_fim = entry_fim.get().strip()

    try:
        if data_inicio and data_fim:
            subprocess.run(
                [sys.executable, SCRIPT_PRINCIPAL, data_inicio, data_fim],
                check=True
            )
        else:
            subprocess.run(
                [sys.executable, SCRIPT_PRINCIPAL],
                check=True
            )

        messagebox.showinfo("Sucesso", "Painel atualizado com sucesso!")

    except subprocess.CalledProcessError:
        messagebox.showerror("Erro", "Erro ao atualizar o painel.\nVerifique o console.")


# ===== INTERFACE =====

janela = tk.Tk()
janela.title("PAINEL COMERCIAL - Atualização")
janela.geometry("500x320")
janela.resizable(False, False)

# Título
titulo = tk.Label(
    janela,
    text="Atualização do Painel Comercial",
    font=("Arial", 16, "bold")
)
titulo.pack(pady=15)

# Frame datas
frame_datas = tk.Frame(janela)
frame_datas.pack(pady=10)

label_inicio = tk.Label(frame_datas, text="Data Inicial (dd/mm/aaaa):")
label_inicio.grid(row=0, column=0, sticky="w")

entry_inicio = tk.Entry(frame_datas, width=20)
entry_inicio.grid(row=1, column=0, padx=10, pady=5)

label_fim = tk.Label(frame_datas, text="Data Final (dd/mm/aaaa):")
label_fim.grid(row=0, column=1, sticky="w")

entry_fim = tk.Entry(frame_datas, width=20)
entry_fim.grid(row=1, column=1, padx=10, pady=5)

# Aviso
aviso = tk.Label(
    janela,
    text="Se deixar as datas vazias, será usada a regra automática.",
    font=("Arial", 9),
    fg="gray"
)
aviso.pack(pady=10)

# Botão atualizar
botao = tk.Button(
    janela,
    text="Atualizar Painel",
    font=("Arial", 12, "bold"),
    bg="#f37021",
    fg="white",
    width=25,
    command=executar_atualizacao
)
botao.pack(pady=20)

# Rodar interface
janela.mainloop()
