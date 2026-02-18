import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import subprocess
import os
import sys

# ===============================
# CONFIGURAÇÕES
# ===============================

SENHA_BASE = "Ondaviperx@"

# ===============================
# VERIFICA SENHA DINÂMICA DO MÊS
# ===============================

def verificar_senha():
    mes = datetime.now().month
    senha_correta = SENHA_BASE + f"{mes:02d}"

    tentativas = 0

    while tentativas < 3:
        senha = tk.simpledialog.askstring(
            "Autenticação",
            "Digite a senha para continuar:",
            show="*"
        )

        if senha == senha_correta:
            return True

        tentativas += 1

    messagebox.showerror("Acesso Bloqueado",
                         "Contate o Emanuel para liberação.")
    return False


# ===============================
# EXECUTA ATUALIZAÇÃO
# ===============================

def executar_atualizacao(data_ini, data_fim):
    try:
        comando = ["python", "python/atualizar_painel_completo.py"]

        if data_ini and data_fim:
            comando.append(data_ini)
            comando.append(data_fim)

        subprocess.run(comando)

        messagebox.showinfo("Sucesso",
                            "Painel atualizado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# ===============================
# INTERFACE PRINCIPAL
# ===============================

def abrir_interface():

    root = tk.Tk()
    root.title("Atualização do Painel Comercial")
    root.geometry("520x350")
    root.resizable(False, False)

    fonte_titulo = ("Arial", 14, "bold")

    tk.Label(root,
             text="Atualização do Painel Comercial",
             font=fonte_titulo).pack(pady=15)

    modo = tk.StringVar(value="auto")

    tk.Radiobutton(root,
                   text="Período Automático (Mês Atual)",
                   variable=modo,
                   value="auto").pack(anchor="w", padx=40)

    tk.Radiobutton(root,
                   text="Período Personalizado",
                   variable=modo,
                   value="manual").pack(anchor="w", padx=40)

    frame_datas = tk.Frame(root)
    frame_datas.pack(pady=15)

    tk.Label(frame_datas, text="Data Inicial (dd/mm/aaaa):").grid(row=0, column=0, sticky="w")
    entry_ini = tk.Entry(frame_datas, width=15)
    entry_ini.grid(row=0, column=1, padx=10)

    tk.Label(frame_datas, text="Data Final (dd/mm/aaaa):").grid(row=1, column=0, sticky="w")
    entry_fim = tk.Entry(frame_datas, width=15)
    entry_fim.grid(row=1, column=1, padx=10)

    def iniciar():
        if modo.get() == "auto":
            executar_atualizacao(None, None)
        else:
            data_ini = entry_ini.get()
            data_fim = entry_fim.get()

            try:
                datetime.strptime(data_ini, "%d/%m/%Y")
                datetime.strptime(data_fim, "%d/%m/%Y")
            except:
                messagebox.showerror("Erro",
                                     "Datas inválidas. Use dd/mm/aaaa.")
                return

            executar_atualizacao(data_ini, data_fim)

    tk.Button(root,
              text="Iniciar Atualização",
              bg="#1976D2",
              fg="white",
              width=25,
              height=2,
              command=iniciar).pack(pady=15)

    tk.Button(root,
              text="Cancelar",
              width=25,
              height=2,
              command=root.destroy).pack()

    root.mainloop()


# ===============================
# EXECUÇÃO
# ===============================

if __name__ == "__main__":
    import tkinter.simpledialog

    if verificar_senha():
        abrir_interface()
