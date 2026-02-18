import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from datetime import datetime
import subprocess
import os

BASE_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def executar_atualizacao(data_inicio=None, data_fim=None):
    try:
        script = os.path.join(BASE_PATH, "python", "atualizar_painel_completo.py")

        if data_inicio and data_fim:
            subprocess.run(
                ["python", script, data_inicio, data_fim],
                check=True
            )
        else:
            subprocess.run(
                ["python", script],
                check=True
            )

        messagebox.showinfo("Sucesso", "Painel atualizado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar:\n{e}")


def atualizar():
    data_inicio = entry_inicio.get().strip()
    data_fim = entry_fim.get().strip()

    if data_inicio and data_fim:
        try:
            datetime.strptime(data_inicio, "%d/%m/%Y")
            datetime.strptime(data_fim, "%d/%m/%Y")
        except:
            messagebox.showerror("Erro", "Formato de data inválido.\nUse DD/MM/AAAA")
            return

        executar_atualizacao(data_inicio, data_fim)
    else:
        executar_atualizacao()


# ==========================
# INTERFACE
# ==========================

root = tk.Tk()
root.title("Painel Comercial - Atualização")
root.geometry("500x350")
root.resizable(False, False)

frame = ttk.Frame(root, padding=20)
frame.pack(expand=True, fill="both")

titulo = ttk.Label(frame, text="Atualização do Painel Comercial", font=("Arial", 16, "bold"))
titulo.pack(pady=10)

sub = ttk.Label(frame, text="Deixe vazio para usar regra automática (01 até hoje)")
sub.pack(pady=5)

lbl_inicio = ttk.Label(frame, text="Data Inicial (DD/MM/AAAA):")
lbl_inicio.pack(pady=(20, 5))

entry_inicio = ttk.Entry(frame, width=20)
entry_inicio.pack()

lbl_fim = ttk.Label(frame, text="Data Final (DD/MM/AAAA):")
lbl_fim.pack(pady=(15, 5))

entry_fim = ttk.Entry(frame, width=20)
entry_fim.pack()

btn = ttk.Button(frame, text="Atualizar Painel", command=atualizar)
btn.pack(pady=30)

root.mainloop()
