from defs import *
import pyautogui as pg
import pyperclip

pg.FAILSAFE = True

#menu
# c = 1 
# c2 = 1
# produtos = []
# while True:
#     peça = []
#     c2 = 1
#     while c2 < 4:
#         codigo = pg.password(F'Digite o {c2}º código da {c}ª peça', mask='', title='Cotação RPA')
#         if codigo == None:
#             break
#         peça.append(codigo)
#         c2 += 1
#     produtos.append(peça)
#     c += 1
#     escolha = pg.confirm('Deseja continuar?', buttons=['Sim', 'Não'], title='Cotação RPA')
#     if escolha == 'Sim':
#         continue
#     else:
#         break

produtos = [['sp271'], ['wo146'], ['sk421', 'ph2966']] #'mb4030', 
# produtos = [['sk421'], ['t36083'], ['vkm4790'], ['ph2966'] , ['mb4030']]#, ['sk423', 'mb4156'], ['40632', '5207110495'], ['880168', '40236', '520423031']]
# produtos = [['c14130'], ['w719/30'], ['sp271']]
codigo = str(produtos[0][0])

#abrir o chrome
pg.hotkey('win', 'r')
time.sleep(0.25)
digitar('chrome')
pg.press('enter')
enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\nova_guia.png')#C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\nova_guia.png
pg.hotkey('win', 'up')  
abrir_site('peca.ai')

#pesquisa peca.ai
enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_peca.ai.png')
clickar_imagem(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\busca_peca.ai.png')
pg.press('tab')
digitar(codigo)
try:
    login = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\logado_peca.ai.png')
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
        sem_resultado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_resultado_peca.ai.png')
    except pg.ImageNotFoundException:
        break
    else:
        pg.press('f5')
        time.sleep(1.5)

try:
    indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_peca.ai.png')
except pg.ImageNotFoundException:
    pg.click(x=560, y=690)
    enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\comprar_peca.ai.png')
    pg.click(x=151, y=651)
    pg.click(x=151, y=651)
    pg.hotkey('ctrl', 'c') 
    unformat()
else:
    pyperclip.copy('Indisponível')

# abrir o excel
time.sleep(0.5)
pg.hotkey('win', 'r')
time.sleep(0.1)
pg.write('excel')
pg.press('enter')
time.sleep(2)
# enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_excel.png')
pg.hotkey('win', 'up')
pg.press('tab')
pg.press('tab')
pg.write('Planilha de Base para Cotacao')   
time.sleep(0.5)
pg.press('enter')
time.sleep(2)

#excel peca.ai
escrever_celula('b2', codigo)
pg.click(x=56, y=181)
pg.click(x=56, y=181)
pg.write('c2')
pg.press('enter')
if str(pyperclip.paste()) == 'Indisponível':
    pg.hotkey('ctrl', 'v')
else:
    pg.write('=')
    pg.hotkey('ctrl', 'v')
    pg.write('+15')
    pg.press('enter')

#pegar nome da peça
alttab()
pg.doubleClick(x=307, y=247)
pg.click(x=307, y=247)
pg.hotkey('ctrl', 'c')
text = pyperclip.paste().split('-')[0]
alttab()
colar_celula('a2', text)

#pesquisa mecanizou
alttab()
pg.hotkey('ctrl', 't')
abrir_site('app.mecanizou.com')
enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou2.png')
pg.press('tab')
pg.write(codigo)
pg.click(x=1162, y=403)
while True:
    try:
        img = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
    except pg.ImageNotFoundException:
        time.sleep(0.5)
        pg.click(x=1162, y=403)
        continue
    else:
        break

try:
    disponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
except pg.ImageNotFoundException:
    pyperclip.copy('Indisponível')
else:
    pg.click(x=487, y=686)
    time.sleep(0.5)
    pg.doubleClick(x=988, y=403)
    pg.hotkey('ctrl', 'c')
    unformat()

#excel mecanizou
alttab()
pg.click(x=56, y=181)
pg.write('d2')
pg.press('enter')
if str(pyperclip.paste()) == 'Indisponível':
    pg.hotkey('ctrl', 'v')
else:
    pg.write('=')
    pg.hotkey('ctrl', 'v')
    pg.write('+7')
    pg.press('enter')  
    
#compel pesquisa
alttab()
pg.hotkey('ctrl', 't')
abrir_site('https://peca.compel.com.br/')
time.sleep(1)
try:
    img = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\acessar_compel.png')
except pg.ImageNotFoundException:
    pass
else:
    pg.click(x=1248, y=366)
enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_compel2.png')
pg.click(x=404, y=417) 
digitar(codigo)
pg.press('enter')
time.sleep(3.5)
try:
    indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_compel.png')
except pg.ImageNotFoundException:
    try:
        pg.click(x=780, y=585)
        time.sleep(1.5)
        estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\disponivel_compel.png')
    except pg.ImageNotFoundException:
        pg.click(x=1125, y=147)
        pyperclip.copy('Indisponível')
    else:
        enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\preco_compel.png')
        pg.doubleClick(x=701, y=339)
        pg.hotkey('ctrl', 'c')
        unformat()
        time.sleep(0.1)
        pg.click(x=1125, y=147)
else:
    pyperclip.copy('Indisponível')

#excel compel
alttab()
pg.click(x=56, y=181)
pg.write('e2')
pg.press('enter')
if str(pyperclip.paste()) == 'Indisponível':
    pg.hotkey('ctrl', 'v')
else:
    pg.write('=')
    pg.hotkey('ctrl', 'v')
    pg.write('+10')
    pg.press('enter')

#dpk pesquisa
alttab()
pg.hotkey('ctrl', 't')
abrir_site('https://www.kdapeca.com.br/login')
time.sleep(3.5)
pg.doubleClick(x=1166, y=709)
time.sleep(1)
pg.press('esc')
time.sleep(0.5)
pg.click(x=314, y=326)
digitar(codigo)
pg.press('enter')
time.sleep(1.5)

pg.click(x=901, y=329)
for c in range(0,3):
    pg.press('down')

try:
    comercializado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\nao_comercializado.png')
except pg.ImageNotFoundException:
    try:
        indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_kdapeca.png')
    except pg.ImageNotFoundException:
        try:
            estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_estoque_kdapeca.png')
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
alttab()
pg.click(x=56, y=181)
pg.write('f2')
pg.press('enter')
if str(pyperclip.paste()) == 'Indisponível':
    pg.hotkey('ctrl', 'v')
else:
    pg.write('=')
    pg.hotkey('ctrl', 'v')
    pg.write('+10')
    pg.press('enter')

#loop a partir do segundo código
celula_peca = 2
for peça in produtos:
    celula_excel = celula_peca
    if peça == produtos[0]:
        celula_peca = 3
        celula_excel = celula_peca
        for codigo in peça:
            if codigo == peça[0]:#nada acontece, apenas altera o indice
                celula_peca = 5
                pass
            elif codigo == peça[-1]:#padrão + altera o indice pra próxima peça              
                alttab()
                pg.hotkey('ctrl', '1')
                pg.hotkey('alt', 'left')
                time.sleep(1)
                for c in range(0, 8):
                    pg.press('tab')
                digitar(codigo)
                pg.press('enter')
                time.sleep(1.5)
                while True:
                    try:
                        sem_resultado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_resultado_peca.ai.png')
                    except pg.ImageNotFoundException:
                        break
                    else:
                        pg.press('f5')
                        time.sleep(1.5)
                try:
                    indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_peca.ai.png')
                except pg.ImageNotFoundException:
                    enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\marcas_peca.ai.png')
                    pg.click(x=560, y=690)
                    enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\comprar_peca.ai.png')
                    pg.click(x=109, y=650)
                    pg.click(x=109, y=650)
                    pg.hotkey('ctrl', 'c') 
                    unformat()
                else:
                    pyperclip.copy('Indisponível')
                
                # excel peca.ai
                alttab()
                escrever_celula('b' + str(celula_excel), codigo)
                pg.click(x=23, y=184)
                pg.click(x=23, y=184)
                pg.write('c' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+15')
                    pg.press('enter')

                #pesquisa mecanizou
                alttab()
                pg.hotkey('ctrl', '2')
                pg.hotkey('alt', 'left')
                enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou2.png')
                pg.press('tab')
                pg.write(codigo)
                pg.click(x=1162, y=403)
                while True:
                    try:
                        img = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
                    except pg.ImageNotFoundException:
                        time.sleep(0.5)
                        pg.click(x=1162, y=403)
                        continue
                    else:
                        break

                try:
                    disponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
                except pg.ImageNotFoundException:
                    pyperclip.copy('Indisponível')
                else:
                    pg.click(x=487, y=686)
                    time.sleep(0.5)
                    pg.doubleClick(x=988, y=403)
                    pg.hotkey('ctrl', 'c')
                    unformat()

                #excel mecanizou
                alttab()
                pg.click(x=56, y=181)
                pg.write('d' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+7')
                    pg.press('enter')

                #compel pesquisa
                alttab()
                pg.hotkey('ctrl', '3')
                pg.doubleClick(x=600, y=367)
                pg.hotkey('ctrl', 'a')
                digitar(codigo)
                pg.press('enter')
                time.sleep(3.5)
                try:
                    indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_compel.png')
                except pg.ImageNotFoundException:
                    try:
                        pg.click(x=780, y=585)
                        time.sleep(1.5)
                        estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\disponivel_compel.png')
                    except pg.ImageNotFoundException:
                        pg.click(x=1125, y=147)
                        pyperclip.copy('Indisponível')
                    else:
                        while True:
                            try:
                                preco = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\preco_compel.png')
                            except pg.ImageNotFoundException:
                                continue
                            else:
                                break
                        pg.doubleClick(x=701, y=339)
                        pg.hotkey('ctrl', 'c')
                        unformat()
                        time.sleep(0.1)
                        pg.click(x=1125, y=147)
                else:
                    pyperclip.copy('Indisponível')

                #excel compel
                alttab()
                pg.click(x=56, y=181)
                pg.write('e' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+10')
                    pg.press('enter')

                #dpk pesquisa
                alttab()
                pg.hotkey('ctrl', '4')
                pg.click(x=15, y=518)
                for c in range(0,3):
                    pg.press('up')
                pg.doubleClick(x=314, y=326)
                pg.hotkey('ctrl', 'a')
                pg.write(codigo)
                pg.press('enter')
                time.sleep(1.5)

                pg.click(x=901, y=329)
                for c in range(0,3):
                    pg.press('down')

                try:
                    comercializado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\nao_comercializado.png')
                except pg.ImageNotFoundException:
                    try:
                        indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_kdapeca.png')
                    except pg.ImageNotFoundException:
                        try:
                            estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_estoque_kdapeca.png')
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
                alttab()
                pg.click(x=56, y=181)
                pg.write('f' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+10')
                    pg.press('enter')

                celula_peca = 5
            else: #padrão
                alttab()
                pg.hotkey('ctrl', '1')
                pg.hotkey('alt', 'left')
                time.sleep(1)
                for c in range(0, 8):
                    pg.press('tab')
                digitar(codigo)
                pg.press('enter')
                time.sleep(1.5)
                while True:
                    try:
                        sem_resultado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_resultado_peca.ai.png')
                    except pg.ImageNotFoundException:
                        break
                    else:
                        pg.press('f5')
                        time.sleep(1.5)
                try:
                    indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_peca.ai.png')
                except pg.ImageNotFoundException:
                    enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\marcas_peca.ai.png')
                    pg.click(x=560, y=690)
                    enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\comprar_peca.ai.png')
                    pg.click(x=109, y=650)
                    pg.click(x=109, y=650)
                    pg.hotkey('ctrl', 'c') 
                    unformat()
                else:
                    pyperclip.copy('Indisponível')
                
                # excel peca.ai
                alttab()
                escrever_celula('b' + str(celula_excel), codigo)
                pg.click(x=23, y=184)
                pg.click(x=23, y=184)
                pg.write('c' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+15')
                    pg.press('enter')

                #pesquisa mecanizou
                alttab()
                pg.hotkey('ctrl', '2')
                pg.hotkey('alt', 'left')
                enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou2.png')
                pg.press('tab')
                pg.write(codigo)
                pg.click(x=1162, y=403)
                while True:
                    try:
                        img = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
                    except pg.ImageNotFoundException:
                        time.sleep(0.5)
                        pg.click(x=1162, y=403)
                        continue
                    else:
                        break

                try:
                    disponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
                except pg.ImageNotFoundException:
                    pyperclip.copy('Indisponível')
                else:
                    pg.click(x=487, y=686)
                    time.sleep(0.5)
                    pg.doubleClick(x=988, y=403)
                    pg.hotkey('ctrl', 'c')
                    unformat()

                #excel mecanizou
                alttab()
                pg.click(x=56, y=181)
                pg.write('d' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+7')
                    pg.press('enter')


                #compel pesquisa
                alttab()
                pg.hotkey('ctrl', '3')
                pg.doubleClick(x=600, y=367)
                pg.hotkey('ctrl', 'a')
                digitar(codigo)
                pg.press('enter')
                time.sleep(3.5)
                try:
                    indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_compel.png')
                except pg.ImageNotFoundException:
                    try:
                        pg.click(x=780, y=585)
                        time.sleep(1.5)
                        estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\disponivel_compel.png')
                    except pg.ImageNotFoundException:
                        pg.click(x=1125, y=147)
                        pyperclip.copy('Indisponível')
                    else:
                        while True:
                            try:
                                preco = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\preco_compel.png')
                            except pg.ImageNotFoundException:
                                continue
                            else:
                                break
                        pg.doubleClick(x=701, y=339)
                        pg.hotkey('ctrl', 'c')
                        unformat()
                        time.sleep(0.1)
                        pg.click(x=1125, y=147)
                else:
                    pyperclip.copy('Indisponível')

                #excel compel
                alttab()
                pg.click(x=56, y=181)
                pg.write('e' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+10')
                    pg.press('enter')

                #dpk pesquisa
                alttab()
                pg.hotkey('ctrl', '4')
                pg.click(x=15, y=518)
                for c in range(0,3):
                    pg.press('up')
                pg.doubleClick(x=314, y=326)
                pg.hotkey('ctrl', 'a')
                pg.write(codigo)
                pg.press('enter')
                time.sleep(1.5)

                pg.click(x=901, y=329)
                for c in range(0,3):
                    pg.press('down')

                try:
                    comercializado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\nao_comercializado.png')
                except pg.ImageNotFoundException:
                    try:
                        indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_kdapeca.png')
                    except pg.ImageNotFoundException:
                        try:
                            estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_estoque_kdapeca.png')
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

                if str(pyperclip.paste()) == codigo:
                    pyperclip.copy('Indisponível')

                #excel dpk
                alttab()
                pg.click(x=56, y=181)
                pg.write('f' + str(celula_excel))
                pg.press('enter')
                if str(pyperclip.paste()) == 'Indisponível':
                    pg.hotkey('ctrl', 'v')
                else:
                    pg.write('=')
                    pg.hotkey('ctrl', 'v')
                    pg.write('+10')
                    pg.press('enter')

                celula_excel += 1
    else:
        for codigo in peça:
            alttab()
            pg.hotkey('ctrl', '1')
            pg.hotkey('alt', 'left')
            time.sleep(1)
            for c in range(0, 8):
                pg.press('tab')
            digitar(codigo) 
            pg.press('enter')
            time.sleep(1.5)
            while True:
                try:
                    sem_resultado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_resultado_peca.ai.png')
                except pg.ImageNotFoundException:
                    break
                else:
                    pg.press('f5')
                    time.sleep(1.5)
            try:
                indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_peca.ai.png')
            except pg.ImageNotFoundException:
                enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\marcas_peca.ai.png')
                pg.click(x=560, y=690)
                enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\comprar_peca.ai.png')
                pg.click(x=109, y=650)
                pg.click(x=109, y=650)
                pg.hotkey('ctrl', 'c') 
                unformat()
            else:   
                pyperclip.copy('Indisponível')
            
            # excel peca.ai
            alttab()
            escrever_celula('b' + str(celula_excel), codigo)
            pg.click(x=23, y=184)
            pg.click(x=23, y=184)
            pg.write('c' + str(celula_excel))
            pg.press('enter')
            if str(pyperclip.paste()) == 'Indisponível':
                pg.hotkey('ctrl', 'v')
            else:
                pg.write('=')
                pg.hotkey('ctrl', 'v')
                pg.write('+15')
                pg.press('enter')

            if codigo == peça[0]:
                #pegar nome da peça
                alttab()
                pg.doubleClick(x=307, y=247)
                pg.click(x=307, y=247)
                pg.hotkey('ctrl', 'c')
                text = pyperclip.paste().split('-')[0]
                alttab()
                colar_celula('a' + str(celula_excel), text)

            #pesquisa mecanizou
            alttab()
            pg.hotkey('ctrl', '2')
            pg.hotkey('alt', 'left')
            enquantonao(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou2.png')
            pg.press('tab')
            pg.write(codigo)
            pg.click(x=1162, y=403)
            while True:
                try:
                    img = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
                except pg.ImageNotFoundException:
                    time.sleep(0.5)
                    pg.click(x=1162, y=403)
                    continue
                else:
                    break

            try:
                disponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\enquantonao_mecanizou.png')
            except pg.ImageNotFoundException:
                pyperclip.copy('Indisponível')
            else:
                pg.click(x=487, y=686)
                time.sleep(0.5)
                pg.doubleClick(x=988, y=403)
                pg.hotkey('ctrl', 'c')
                unformat()

            #excel mecanizou
            alttab()
            pg.click(x=56, y=181)
            pg.write('d' + str(celula_excel))
            pg.press('enter')
            if str(pyperclip.paste()) == 'Indisponível':
                pg.hotkey('ctrl', 'v')
            else:
                pg.write('=')
                pg.hotkey('ctrl', 'v')
                pg.write('+7')
                pg.press('enter')

            #compel pesquisa
            alttab()
            pg.hotkey('ctrl', '3')
            pg.doubleClick(x=600, y=367)
            pg.hotkey('ctrl', 'a')
            digitar(codigo)
            pg.press('enter')
            time.sleep(3.5)
            try:
                indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_compel.png')
            except pg.ImageNotFoundException:
                try:
                    pg.click(x=780, y=585)
                    time.sleep(1.5)
                    estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\disponivel_compel.png')
                except pg.ImageNotFoundException:
                    pg.click(x=1125, y=147)
                    pyperclip.copy('Indisponível')
                else:
                    while True:
                        try:
                            preco = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\preco_compel.png')
                        except pg.ImageNotFoundException:
                            continue
                        else:
                            break
                    pg.doubleClick(x=701, y=339)
                    pg.hotkey('ctrl', 'c')
                    unformat()
                    time.sleep(0.1)
                    pg.click(x=1125, y=147)
            else:
                pyperclip.copy('Indisponível')

            #excel compel
            alttab()
            pg.click(x=56, y=181)
            pg.write('e' + str(celula_excel))
            pg.press('enter')
            if str(pyperclip.paste()) == 'Indisponível':
                pg.hotkey('ctrl', 'v')
            else:
                pg.write('=')
                pg.hotkey('ctrl', 'v')
                pg.write('+10')
                pg.press('enter')

            #dpk pesquisa
            alttab()
            pg.hotkey('ctrl', '4')
            pg.click(x=15, y=518)
            for c in range(0,3):
                pg.press('up')
            pg.doubleClick(x=314, y=326)
            pg.hotkey('ctrl', 'a')
            pg.write(codigo)
            pg.press('enter')
            time.sleep(1.5)

            pg.click(x=901, y=329)
            for c in range(0,3):
                pg.press('down')

            try:
                comercializado = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\nao_comercializado.png')
            except pg.ImageNotFoundException:
                try:
                    indisponivel = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\indisponivel_kdapeca.png')
                except pg.ImageNotFoundException:
                    try:
                        estoque = pg.locateOnScreen(r'C:\Users\Dell\OneDrive\Documentos\GitHub\hawks-automations\imagens\sem_estoque_kdapeca.png')
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
            alttab()
            pg.click(x=56, y=181)
            pg.write('f' + str(celula_excel))
            pg.press('enter')
            if str(pyperclip.paste()) == 'Indisponível':
                pg.hotkey('ctrl', 'v')
            else:
                pg.write('=')
                pg.hotkey('ctrl', 'v')
                pg.write('+10')
                pg.press('enter')
            #não mecher
            celula_excel += 1
    
    if peça != produtos[0]:
        celula_peca += 3
