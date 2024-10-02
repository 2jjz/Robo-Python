from defs import *
import pyautogui as pg
import pyperclip
import time

for c in range(0,2):
    with pg.hold('alt'):
        time.sleep(0.2)
        pg.press('tab')
        time.sleep(0.2)
        pg.press('tab')
        time.sleep(0.2)
    print()

pg.FAILSAFE = True
# produtos = [['sk421'], ['t36083'], ['vkm4790'], ['ph2966'] , ['mb4030']]#, ['sk423', 'mb4156'], ['40632', '5207110495'], ['880168', '40236', '520423031']]
produtos = [['sk421', 'sp271', 'wo1', 't36083', 'vkm4790', 'ph2966']]
codigo = produtos[0][0]



celula_excel = 3
#test second processos
for peca in produtos:
    for codigo in peca:

        #processo
        #pesquisa rmp
        alttab()
        pg.hotkey('ctrl', '5')
        pg.click(x=285, y=138)
        pg.hotkey('ctrl', 'a')
        pg.write(codigo)
        time.sleep(0.5)
        pg.press('enter')
        choice = pg.confirm('Escolha a peça e depois clique no título.', buttons=['Ok', 'Peça Indisponível'], title='Cotação RPA')

        if choice == 'Ok':
            clickar_imagem(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sifrao_rpm.png', 3)
            pg.hotkey('ctrl', 'c')
            pyperclip.copy(pyperclip.paste()[2:])
        else:
            pyperclip.copy('Indisponível')
        print(pyperclip.paste())

        #excel rmp
        alttab()
        pg.click(x=56, y=181)
        pg.click(x=56, y=181)
        pg.write('g' + str(celula_excel))
        pg.press('enter')
        if str(pyperclip.paste()) == 'Indisponível':
            pg.hotkey('ctrl', 'v')
        else:
            pg.write('=')
            pg.hotkey('ctrl', 'v')
            pg.write('+40')
            pg.press('enter')

        celula_excel += 1