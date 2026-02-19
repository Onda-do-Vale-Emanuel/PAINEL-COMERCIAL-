import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
import subprocess
import atualizar_painel_completo as painel  # usa o seu script atual


# ======================================
# CONFIGURAÇÕES BÁSICAS
# ======================================
BASE_DIR = Path(__file__).resolve().parent.parent
PY_DIR = BASE_DIR / "python"

# Arquivos para controle da senha
CONTADOR_PATH = PY_DIR / ".contador_exec"
MES_PATH = PY_DIR / ".mes_exec"

# Quantidade de execuções permitidas sem pedir senha
MAX_EXEC_SEM_SENHA = 30


# ======================================
# FUNÇÕES DE SENHA
# ======================================
def senha_mes_atual() -> str:
    """Gera a senha do mês: Ondaviperx@MM"""
    mes = datetime.now().month
    return f"Ondaviperx@{mes:02d}"


def ler_contador() -> int:
    """Lê o contador de execuções e reseta no início de cada mês."""
    try:
        mes_atual = f"{datetime.now().year}-{datetime.now().month:02d}"

        # Atualiza / cria arquivo com mês atual
        if MES_PATH.exists():
            mes_salvo = MES_PATH.read_text(encoding="utf-8").strip()
            if mes_salvo != mes_atual:
                # Novo mês → zera contador
                CONTADOR_PATH.write_text("0", encoding="utf-8")
                MES_PATH.write_text(mes_atual, encoding="utf-8")
        else:
            MES_PATH.write_text(mes_atual, encoding="utf-8")

        # Lê contador
        if CONTADOR_PATH.exists():
            valor = CONTADOR_PATH.read_text(encoding="utf-8").strip()
            return int(valor or "0")
        else:
            CONTADOR_PATH.write_text("0", encoding="utf-8")
            return 0
    except:
        # Se der erro por qualquer motivo, não trava o sistema
        return 0


def salvar_contador(valor: int) -> None:
    try:
        CONTADOR_PATH.write_text(str(valor), encoding="utf-8")
    except:
        # Não quebra o programa se não conseguir salvar
        pass


def pedir_senha_se_necessario(root: tk.Tk) -> bool:
    """
    Verifica contador. Se já passou do limite, pede a senha.
    Retorna True se pode continuar, False se deve abortar.
    """
    contador = ler_contador()

    # Ainda dentro da cota → não pede senha
    if contador < MAX_EXEC_SEM_SENHA:
        salvar_contador(contador + 1)
        return True

    senha_correta = senha_mes_atual()
    tentativas = 0

    while tentativas < 3:
        dlg = tk.Toplevel(root)
        dlg.title("Autenticação necessária")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.transient(root)

        # Centralizar a janelinha
        dlg.update_idletasks()
        largura = 350
        altura = 150
        x = root.winfo_rootx() + (root.winfo_width() // 2) - (largura // 2)
        y = root.winfo_rooty() + (root.winfo_height() // 2) - (altura // 2)
        dlg.geometry(f"{largura}x{altura}+{x}+{y}")

        tk.Label(
            dlg,
            text="Digite a senha para continuar:",
            font=("Segoe UI", 10)
        ).pack(padx=20, pady=(20, 5))

        senha_var = tk.StringVar()
        entry = tk.Entry(dlg, textvariable=senha_var, show="*", width=30)
        entry.pack(padx=20, pady=5)
        entry.focus_set()

        resultado = {"ok": False}

        def confirmar():
            if senha_var.get() == senha_correta:
                resultado["ok"] = True
                dlg.destroy()
            else:
                messagebox.showerror("Senha incorreta", "Senha inválida. Tente novamente.")
                entry.delete(0, tk.END)
                entry.focus_set()

        def cancelar():
            dlg.destroy()

        frame_btn = tk.Frame(dlg)
        frame_btn.pack(pady=(10, 15))

        ttk.Button(frame_btn, text="OK", width=10, command=confirmar).pack(side="left", padx=5)
        ttk.Button(frame_btn, text="Cancelar", width=10, command=cancelar).pack(side="left", padx=5)

        dlg.bind("<Return>", lambda e: confirmar())
        dlg.bind("<Escape>", lambda e: cancelar())

        root.wait_window(dlg)

        if resultado["ok"]:
            # Zera contador → libera mais 30 execuções
            salvar_contador(0)
            return True

        tentativas += 1

    messagebox.showerror("Acesso bloqueado", "Contate o Emanuel para liberação.")
    return False


# ======================================
# FUNÇÃO PARA DAR GIT PUSH
# ======================================
def enviar_ao_github():
    """
    Roda git add / git commit / git push na pasta do projeto.
    Se der erro, mostra um aviso mas não considera a atualização perdida.
    """
    try:
        # git add .
        subprocess.run(
            ["git", "add", "."],
            cwd=str(BASE_DIR),
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # git commit -m "Atualização painel YYYY-MM-DD HH:MM"
        msg = "Atualização painel " + datetime.now().strftime("%Y-%m-%d %H:%M")
        subprocess.run(
            ["git", "commit", "-m", msg],
            cwd=str(BASE_DIR),
            check=False,  # pode não ter nada pra commitar
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # git push origin main
        resultado = subprocess.run(
            ["git", "push", "origin", "main"],
            cwd=str(BASE_DIR),
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        print(resultado.stdout)
        return True

    except Exception as e:
        messagebox.showwarning(
            "Aviso",
            "Painel atualizado localmente, mas houve erro ao enviar para o GitHub:\n\n"
            f"{e}"
        )
        return False


# ======================================
# FUNÇÃO PRINCIPAL DA INTERFACE
# ======================================
def criar_interface():
    root = tk.Tk()
    root.title("Atualização do Painel Comercial")
    root.geometry("600x320")
    root.resizable(False, False)

    # Deixa com cara mais moderna (Windows)
    try:
        root.iconbitmap(default="")  # se quiser colocar um ícone .ico depois
    except:
        pass

    # Título
    titulo = tk.Label(
        root,
        text="Atualização do Painel Comercial",
        font=("Segoe UI", 16, "bold")
    )
    titulo.pack(pady=(20, 10))

    # Frame dos modos
    frame_modo = tk.Frame(root)
    frame_modo.pack(pady=(5, 10), fill="x", padx=40)

    modo_var = tk.IntVar(value=1)  # 1 = automático, 2 = personalizado

    rb_auto = tk.Radiobutton(
        frame_modo,
        text=" Período Automático (Mês Atual)",
        variable=modo_var,
        value=1,
        font=("Segoe UI", 10)
    )
    rb_auto.grid(row=0, column=0, sticky="w", pady=2)

    rb_perso = tk.Radiobutton(
        frame_modo,
        text=" Período Personalizado",
        variable=modo_var,
        value=2,
        font=("Segoe UI", 10)
    )
    rb_perso.grid(row=1, column=0, sticky="w", pady=2)

    # Frame das datas
    frame_datas = tk.Frame(root)
    frame_datas.pack(pady=(10, 10), padx=40, fill="x")

    tk.Label(frame_datas, text="Data Inicial:", font=("Segoe UI", 10)).grid(
        row=0, column=0, sticky="e", pady=5
    )
    tk.Label(frame_datas, text="Data Final:", font=("Segoe UI", 10)).grid(
        row=1, column=0, sticky="e", pady=5
    )

    data_ini_var = tk.StringVar()
    data_fim_var = tk.StringVar()

    entry_ini = tk.Entry(frame_datas, textvariable=data_ini_var, width=15, font=("Segoe UI", 10))
    entry_fim = tk.Entry(frame_datas, textvariable=data_fim_var, width=15, font=("Segoe UI", 10))

    entry_ini.grid(row=0, column=1, sticky="w", padx=10, pady=5)
    entry_fim.grid(row=1, column=1, sticky="w", padx=10, pady=5)

    # Linha de separação
    separador = ttk.Separator(root, orient="horizontal")
    separador.pack(fill="x", padx=40, pady=(10, 15))

    # Frame dos botões
    frame_botoes = tk.Frame(root)
    frame_botoes.pack(pady=(5, 15))

    def habilitar_datas():
        """Habilita/desabilita as caixas de data conforme o modo."""
        if modo_var.get() == 1:
            entry_ini.configure(state="disabled")
            entry_fim.configure(state="disabled")
        else:
            entry_ini.configure(state="normal")
            entry_fim.configure(state="normal")

    habilitar_datas()  # aplica no início

    rb_auto.configure(command=habilitar_datas)
    rb_perso.configure(command=habilitar_datas)

    def iniciar_atualizacao():
        # 1) Verificar senha / contador
        if not pedir_senha_se_necessario(root):
            return

        # 2) Rodar atualização python
        try:
            if modo_var.get() == 1:
                # Modo automático: mês atual
                painel.main()
            else:
                ini = data_ini_var.get().strip()
                fim = data_fim_var.get().strip()

                if not ini or not fim:
                    messagebox.showerror(
                        "Datas inválidas",
                        "Preencha a data inicial e a data final\nou escolha o modo automático."
                    )
                    return

                # valida formato
                try:
                    datetime.strptime(ini, "%d/%m/%Y")
                    datetime.strptime(fim, "%d/%m/%Y")
                except ValueError:
                    messagebox.showerror(
                        "Formato de data inválido",
                        "Use o formato dd/mm/aaaa.\nExemplo: 01/02/2026"
                    )
                    return

                painel.main(ini, fim)

        except Exception as e:
            messagebox.showerror(
                "Erro na atualização",
                f"Ocorreu um erro ao atualizar o painel:\n\n{e}"
            )
            return

        # 3) Enviar para GitHub
        sucesso_push = enviar_ao_github()

        if sucesso_push:
            messagebox.showinfo(
                "Sucesso",
                "Painel atualizado e enviado para o GitHub com sucesso!"
            )
        else:
            messagebox.showinfo(
                "Sucesso parcial",
                "Painel atualizado localmente.\n\n"
                "Houve um problema ao enviar para o GitHub.\n"
                "Verifique a conexão ou as credenciais e tente novamente."
            )

    btn_iniciar = tk.Button(
        frame_botoes,
        text="Iniciar Atualização",
        font=("Segoe UI", 11, "bold"),
        bg="#0078D7",
        fg="white",
        activebackground="#005A9E",
        activeforeground="white",
        width=20,
        command=iniciar_atualizacao
    )
    btn_iniciar.grid(row=0, column=0, padx=10)

    btn_cancelar = tk.Button(
        frame_botoes,
        text="Cancelar",
        font=("Segoe UI", 11),
        width=12,
        command=root.destroy
    )
    btn_cancelar.grid(row=0, column=1, padx=10)

    root.mainloop()


if __name__ == "__main__":
    criar_interface()
