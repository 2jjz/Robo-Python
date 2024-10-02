from defs import *
import pyautogui as pg
import pyperclip
import time

# TESTE DE MODULARIZAÇÃO SEGUNDA RODADA

produtos = [['w719/30','sp271'], ['T36083', 'sp271', 'sk421'], ['vkm4790']]

#TAB SET
for c in range(0,2):
    with pg.hold('alt'):
        time.sleep(0.2)
        pg.press('tab')
        time.sleep(0.2)
        pg.press('tab')
        time.sleep(0.2)
    print()

pegar_nome_rmp(False, 'a8')