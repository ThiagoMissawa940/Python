import pyautogui
import time
import pandas as pd
import numpy
import openpyxl



tabela = pd.read_csv('./Relatorio_SBTI_convertido_iso.csv', encoding='ISO-8859-1', delimiter=';')
pyautogui.PAUSE = 1.5


time.sleep(5)
for linha in tabela.index:

    #primeira filtragem
    pyautogui.click(x=121, y=376)
    pyautogui.click(x=1581, y=261)
    pyautogui.click(x=1013, y=626)
    pyautogui.write(str(tabela.loc[linha, "E-mail"]))
    pyautogui.press('enter')
    
    # Cadastrar produto
    pyautogui.click(x=1208, y=702)
    pyautogui.click(x=1672, y=263)
    pyautogui.click(x=585, y=429)

    pyautogui.write(str(tabela.loc[linha, "Nome"]))
    pyautogui.press('tab')
    pyautogui.write(str(tabela.loc[linha, "E-mail"]))
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write('godental')
    pyautogui.press('tab')
    pyautogui.write('godental')
    pyautogui.click(x=1026, y=896)

    # segunda filtragem
    pyautogui.click(x=121, y=376)
    pyautogui.click(x=1580, y=262)
    pyautogui.click(x=1079, y=643)
    pyautogui.write(str(tabela.loc[linha, "E-mail"]))
    pyautogui.press('enter')
    pyautogui.click(x=811, y=403)

    # registrar plano
    pyautogui.click(x=1917, y=863)
    pyautogui.click(x=1673, y=959)
    pyautogui.click(x=1101, y=553)
    pyautogui.write('SBTI')
    pyautogui.click(x=1209, y=863)


    pyautogui.scroll(600)
    pyautogui.click()
