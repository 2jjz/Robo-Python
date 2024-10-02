import pyautogui as pg
import time
import pyperclip

def enquantonao(imagem):
    while True:
        try:
            img = pg.locateOnScreen(imagem)
        except pg.ImageNotFoundException:
            continue
        else:
            break

def clickar_imagem(img, vezes=1):
    x, y, largura, altura = pg.locateOnScreen(img, grayscale=True)
    x = x + largura / 2
    y = y + altura / 2
    for c in range(0, vezes):
        pg.click(x, y)
    

def preencher_campo(info):
    pyperclip.copy(info)
    pg.hotkey('ctrl', 'a')
    pg.hotkey('ctrl', 'v')
    pg.press('tab')

def digitar(txt):
    pyperclip.copy(txt)
    pg.hotkey('ctrl', 'v')

def escrever_celula(cell, txt):
    pg.click(x=23, y=184)
    pg.click(x=23, y=184)
    pg.write(cell)
    time.sleep(0.25)
    pg.press('enter')
    time.sleep(0.25)
    pg.write(txt)
    time.sleep(0.25)
    pg.press('enter')
    time.sleep(0.25)

def colar_celula(cell, txt):
    pg.click(x=23, y=184)
    pg.click(x=23, y=184)
    pg.write(cell)
    time.sleep(0.25)
    pg.press('enter')
    time.sleep(0.25)
    pg.click(x=259, y=183)
    time.sleep(0.25)
    pyperclip.copy(txt)
    pg.write(txt)
    # pg.hotkey('ctrl' ,'v')
    time.sleep(0.25)
    pg.press('enter')
    time.sleep(0.25)

def abrir_site(link): 
    pg.click(x=182, y=67)
    pg.click(x=182, y=67)
    pg.hotkey('ctrl', 'a')
    pyperclip.copy(link)
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    time.sleep(3)

def pegar_nome_peca_ai(tem_nome, celula):
    if tem_nome == True:
        return True
    else:
        alttab()
        pg.doubleClick(x=307, y=247)
        pg.click(x=307, y=247)
        pg.hotkey('ctrl', 'c')
        text = pyperclip.paste().split('-')[0]
        text = text.split(' ')
        new_text = ''
        for c in range(0, 2):
            text.pop()
        for i in text:
            new_text += i
            new_text += ' '
        text = new_text
        alttab()
        colar_celula(celula, text)
        return True
    
def pegar_nome_mecanizou(tem_nome, celula):
    if tem_nome == True:
        return True
    else:
        alttab()
        pg.doubleClick(x=966, y=333)
        pg.click(x=966, y=333)
        pg.hotkey('ctrl', 'c')
        text = pyperclip.paste()[:-1]
        pyperclip.copy(text)
        alttab()
        colar_celula(celula, text)
        return True

def pegar_nome_compel(tem_nome, celula):
    if tem_nome == True:
        return True
    else:
        # alttab()
        pg.doubleClick(x=645, y=210)
        pg.click(x=645, y=210)
        pg.hotkey('ctrl', 'c')
        text = pyperclip.paste()
        new_text = ''
        text = text.split(' ')
        for c in range(0, 2):
            text.pop()
        for c in text:
            new_text += c
            new_text += ' '
        text = new_text
        alttab()
        colar_celula(celula, text)
        alttab()
        return True

def pegar_nome_dpk(tem_nome, celula):
    if tem_nome == True:
        return True
    else:
        alttab()
        pg.click(x=125, y=575)
        time.sleep(1.2)
        pg.doubleClick(x=94, y=219)
        pg.click(x=94, y=219)
        pg.hotkey('ctrl', 'c')
        pg.press('esc')
        text = pyperclip.paste()
        text = text.split(' ')
        new_text = ''
        for c in range(0, 3):
            text.pop()
        for i in text:
            new_text += i
            new_text += ' '
        text = new_text
        pyperclip.copy(text)
        alttab()
        colar_celula(celula, text)
        return True
    
def pegar_nome_rmp(tem_nome, celula):
    if tem_nome == True:
        return True
    else:
        alttab()
        pg.doubleClick(x=610, y=351)
        pg.click(x=610, y=351)
        pg.hotkey('ctrl', 'c')
        text = pyperclip.paste()
        text = text.split('-')[0]
        pyperclip.copy(text)
        alttab()
        colar_celula(celula, text)
        return True

def unformat():
    pg.press('win')
    time.sleep(0.35)
    pg.hotkey('ctrl', 'v')
    time.sleep(0.35)
    pg.hotkey('ctrl', 'a')
    time.sleep(0.35)
    pg.hotkey('ctrl', 'c')
    time.sleep(0.35)
    pg.press('win')
    pg.press('win')

def alttab():
    with pg.hold('alt'):
        time.sleep(0.1)
        pg.press('tab')
    time.sleep(0.5)

def zoom():
    with pg.hold('ctrl'):
        time.sleep(0.5)
        pg.scroll(100)
        time.sleep(0.5)
        pg.scroll(100)
        time.sleep(0.5)
        pg.scroll(100)

#-------

#abrir o chrome
def abrir_chrome():
    pg.hotkey('win', 'r')
    time.sleep(0.25)
    digitar('chrome')
    pg.press('enter')
    enquantonao(r'imagens/nova_guia.png')
    pg.hotkey('win', 'up')

#pesquisa peca.ai
def pesquisa_peca_ai(codigo):
    enquantonao(r'imagens/enquantonao_peca.ai.png')
    clickar_imagem(r'imagens/busca_peca.ai.png')
    pg.press('tab')
    digitar(codigo)
    try:
        login = pg.locateOnScreen(r'imagens/logado_peca.ai.png')
    except pg.ImageNotFoundException:
        pg.press('enter')
        time.sleep(1)
        for c in range (0, 3):
            pg.press('tab')
        pg.press('enter')
    else:
        pg.press('enter')
    time.sleep(1.5)
    while True:
        try:
            sem_resultado = pg.locateOnScreen(r'imagens/sem_resultado_peca.ai.png')
        except pg.ImageNotFoundException:
            break
        else:
            pg.press('f5')
            time.sleep(1.5)
    try:
        indisponivel = pg.locateOnScreen(r'imagens/indisponivel_peca.ai.png')
    except pg.ImageNotFoundException:
        pg.click(x=560, y=690)
        enquantonao(r'imagens/comprar_peca.ai.png')
        pg.click(x=151, y=651)
        pg.click(x=151, y=651)
        pg.hotkey('ctrl', 'c') 
        unformat()
    else:
        pyperclip.copy('Indisponível')

# abrir o excel
def abrir_excel():
    time.sleep(0.5)
    pg.hotkey('win', 'r')
    pg.write('excel')
    pg.press('enter')
    enquantonao(r'imagens/enquantonao_excel2.png')
    pg.hotkey('win', 'up')
    pg.press('tab')
    pg.press('tab')
    pg.write('Planilha de Base para Cotacao')   
    time.sleep(0.5)
    pg.press('enter')
    time.sleep(2)

#excel peca.ai
def excel_peca_ai(cell, codigo):
    escrever_celula('b' + str(cell), codigo)
    pg.click(x=56, y=181)
    pg.click(x=56, y=181)
    pg.write('c' + str(cell))
    pg.press('enter')
    if str(pyperclip.paste()) == 'Indisponível':
        pg.hotkey('ctrl', 'v')
    else:
        pg.write('=')
        pg.hotkey('ctrl', 'v')
        pg.write('+15')
        pg.press('enter')

def pegar_nome_peca(cell):
    #pegar nome da peça
    pg.doubleClick(x=307, y=247)
    pg.click(x=307, y=247)
    pg.hotkey('ctrl', 'c')
    text = pyperclip.paste().split('-')[0]
    alttab()
    colar_celula(cell, text)

def pesquisa_mecanizou(codigo):
    #pesquisa mecanizou
    enquantonao(r'imagens/enquantonao_mecanizou2.png')
    pg.press('tab')
    digitar(codigo)
    pg.press('enter')
    enquantonao(r'imagens/marcas_mecanizou.png')
    try:
        disponivel = pg.locateOnScreen(r'imagens/enquantonao_mecanizou.png')
    except pg.ImageNotFoundException:
        pyperclip.copy('Indisponível')
    else:
        pg.click(x=500, y=491)
        time.sleep(0.5)
        pg.doubleClick(x=988, y=403)
        pg.hotkey('ctrl', 'c')
        unformat()

def excel_mecanizou(cell):
    #excel mecanizou
    pg.click(x=56, y=181)
    pg.write('d' + str(cell))
    pg.press('enter')
    if str(pyperclip.paste()) == 'Indisponível':
        pg.hotkey('ctrl', 'v')
    else:
        pg.write('=')
        pg.hotkey('ctrl', 'v')
        pg.write('+7')
        pg.press('enter')

#compel pesquisa
def pesquisa_compel(codigo):
    try:
        img = pg.locateOnScreen(r'imagens/acessar_compel.png')
    except pg.ImageNotFoundException:
        pass
    else:
        pg.click(x=1248, y=366)
    enquantonao(r'imagens/enquantonao_compel2.png')
    pg.click(x=404, y=417) 
    digitar(codigo)
    pg.press('enter')
    time.sleep(3.5)
    try:
        indisponivel = pg.locateOnScreen(r'imagens/indisponivel_compel.png')
    except pg.ImageNotFoundException:
        try:
            pg.click(x=780, y=585)
            time.sleep(1.5)
            estoque = pg.locateOnScreen(r'imagens/disponivel_compel.png')
        except pg.ImageNotFoundException:
            pg.click(x=1125, y=147)
            pyperclip.copy('Indisponível')
        else:
            enquantonao(r'imagens/preco_compel.png')
            pg.doubleClick(x=701, y=339)
            pg.hotkey('ctrl', 'c')
            unformat()
            time.sleep(0.1)
            pg.click(x=1125, y=147)
    else:
        pyperclip.copy('Indisponível')

#excel compel
def excel_compel(cell):
    pg.click(x=56, y=181)
    pg.write('e' + str(cell))
    pg.press('enter')
    if str(pyperclip.paste()) == 'Indisponível':
        pg.hotkey('ctrl', 'v')
    else:
        pg.write('=')
        pg.hotkey('ctrl', 'v')
        pg.write('+10')
        pg.press('enter')

#dpk pesquisa
def pesquisa_kdapeca(codigo):    
    pg.click(x=314, y=326)
    pg.hotkey('ctrl', 'a')
    digitar(codigo)
    pg.press('enter')
    time.sleep(1.5)
    pg.click(x=901, y=329)
    for c in range(0,3):
        pg.press('down')
    try:
        comercializado = pg.locateOnScreen(r'imagens/nao_comercializado.png')
    except pg.ImageNotFoundException:
        try:
            indisponivel = pg.locateOnScreen(r'imagens/indisponivel_kdapeca.png')
        except pg.ImageNotFoundException:
            try:
                estoque = pg.locateOnScreen(r'imagens/sem_estoque_kdapeca.png')
            except pg.ImageNotFoundException:
                pg.doubleClick(x=788, y=705)
                pg.hotkey('ctrl', 'c')
                unformat()
            else:
                pyperclip.copy('Indisponível')    
        else:
            pyperclip.copy('Indisponível')
    else:
        pyperclip.copy('Indisponível')

#excel dpk
def excel_kdapeca(cell):
    pg.click(x=56, y=181)
    pg.write('f' + str(cell))
    pg.press('enter')
    if str(pyperclip.paste()) == 'Indisponível':
        pg.hotkey('ctrl', 'v')
    else:
        pg.write('=')
        pg.hotkey('ctrl', 'v')
        pg.write('+10')
        pg.press('enter')
