from defs import *
import pyautogui as pg
import pyperclip
import time

# alttab()
while True:
    try:
        time.sleep(1)
        img = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\marca_mecanizou.png')
    except pg.ImageNotFoundException:
        print('não achei')
        continue
    else:
        print('achei')
        # break
