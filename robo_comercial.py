# ===============================================================
# ROBO COMERCIAL - INFOBOX (MODO VISÍVEL - ESTÁVEL)
# ===============================================================

import pyautogui
import time
import subprocess
from datetime import datetime
import pandas as pd
import pyperclip
import pygetwindow as gw

# ===============================================================
# CONFIG
# ===============================================================

ERP_PATH = r"I:\InfoBoxCS.exe"
EXCEL_PATH = r"D:\USUARIOS\ADM05\Documents\dashboard_tv\excel\PEDIDOS_2026.xlsx"

# MAIS LENTO PRA VOCÊ VER
pyautogui.PAUSE = 0.8


# ===============================================================
# FORÇAR FOCO + MAXIMIZAR
# ===============================================================

def focar_e_maximizar():
    janelas = gw.getAllWindows()

    for janela in janelas:
        try:
            titulo = janela.title.lower()
            if "info" in titulo or "box" in titulo:
                
                if janela.isMinimized:
                    janela.restore()

                janela.activate()
                time.sleep(1)

                try:
                    janela.maximize()
                except:
                    pass

                time.sleep(1)

                # clique no meio da tela
                x, y = janela.center
                pyautogui.click(x, y)

                return True
        except:
            pass

    print("⚠️ Não achou janela, tentando ALT+TAB")

    # fallback
    pyautogui.keyDown("alt")
    pyautogui.press("tab")
    pyautogui.keyUp("alt")

    time.sleep(1)
    pyautogui.click(600, 400)

    return True


# ===============================================================
# ABRIR ERP
# ===============================================================

def abrir_erp():
    subprocess.Popen(ERP_PATH)
    time.sleep(6)

    focar_e_maximizar()


# ===============================================================
# LOGIN VISÍVEL
# ===============================================================

def login():
    focar_e_maximizar()

    print("➡️ Digitando login...")
    time.sleep(2)

    pyautogui.write("onda do vale")
    pyautogui.press("tab")

    pyautogui.write("EMANUEL")
    pyautogui.press("tab")

    pyautogui.write("EMANUEL")
    pyautogui.press("enter")

    time.sleep(5)

    pyautogui.press("n")
    time.sleep(2)


# ===============================================================
# NAVEGAÇÃO
# ===============================================================

def abrir_relatorio():
    focar_e_maximizar()

    print("➡️ Abrindo pedidos...")

    pyautogui.press("alt")
    time.sleep(1)

    pyautogui.press("m")
    time.sleep(1)

    pyautogui.press("v")
    time.sleep(1)

    pyautogui.press("p")

    time.sleep(5)


# ===============================================================
# DEFINIR PERÍODO
# ===============================================================

def definir_periodo():
    focar_e_maximizar()

    print("➡️ Definindo período...")

    hoje = datetime.now()
    inicio = hoje.replace(day=1)

    data_ini = inicio.strftime("%d%m%Y")
    data_fim = hoje.strftime("%d%m%Y")

    pyautogui.write(data_ini)
    pyautogui.press("tab")
    pyautogui.write(data_fim)

    pyautogui.press("tab")
    pyautogui.press("space")

    pyautogui.press("enter")

    time.sleep(10)


# ===============================================================
# COPIAR DADOS
# ===============================================================

def copiar_dados():
    focar_e_maximizar()

    print("➡️ Copiando dados...")

    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("ctrl", "c")

    time.sleep(2)

    return pyperclip.paste()


# ===============================================================
# ATUALIZAR EXCEL
# ===============================================================

def atualizar_excel(texto):
    print("➡️ Atualizando Excel...")

    linhas = texto.split("\n")
    dados = [linha.split("\t") for linha in linhas if linha.strip()]

    df = pd.DataFrame(dados)

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        df.to_excel(writer, index=False, header=False, startrow=1)


# ===============================================================
# ATUALIZAR PAINEL
# ===============================================================

def atualizar_painel():
    print("➡️ Atualizando painel...")
    import atualizar_painel_completo as painel
    painel.main()


# ===============================================================
# EXECUÇÃO COMPLETA
# ===============================================================

def executar_robo():
    abrir_erp()
    login()
    abrir_relatorio()
    definir_periodo()

    dados = copiar_dados()
    atualizar_excel(dados)

    atualizar_painel()


# ===============================================================

if __name__ == "__main__":
    executar_robo()