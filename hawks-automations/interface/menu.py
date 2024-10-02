import pymsgbox as pg

def menu():
    pg.CANCEL_TEXT = 'Finalizar'
    pg.OK_TEXT = 'Próximo'
    cont_peca = 1 
    cont_code = 1
    produtos = []
    while True:

        #adicionar peças
        while True:
            peça = []
            cont_code = 1
            while cont_code < 4:
                codigo = pg.password(f'Digite o {cont_code}º código da {cont_peca}ª peça', mask='', title='Cotação RPA') #FINISH
                if codigo == None:
                    break
                peça.append(codigo)
                cont_code += 1
            produtos.append(peça)
            cont_peca += 1
            escolha = pg.confirm('Deseja continuar a adicionar peças?', buttons=['Sim', 'Não'], title='Cotação RPA')
            if escolha == 'Sim':
                continue
            else:
                break
        
        #formatação de texto
        txt = ''
        pg.CANCEL_TEXT = 'Cancelar'
        pg.OK_TEXT = 'Ok'
        
        while True:

            # formatação da lista
            for indice, peca in enumerate(produtos):
                if len(peca) == 0:
                    produtos.pop(indice)

            #formatação de texto
            txt = ''
            for indice, peca in enumerate(produtos):
                txt += f'''Peça {indice + 1} - {peca}
    '''
            escolha = pg.confirm(f'''Peças cotadas:
    {txt}''', buttons=['Continuar', 'Adicionar mais peças', 'Excluir Peça', 'Corrigir'], title='Cotação RPA')
            
            #escolhas
            if escolha == 'Corrigir':
                while True:
                    #pegar peca a ser corrigida
                    try:
                        peca = int(pg.password(f'''Peças cotadas:
    {txt}
    Qual peça você deseja corrigir os códigos? (Digite apenas o número)''', title='Cotação RPA', mask='')) - 1
                        codigo_corrigido = produtos[peca]
                    except ValueError:
                        pg.alert('Insira um valor numérico!')
                        continue
                    except IndexError:
                        pg.alert(f'Não há peça {peca + 1}!')
                        continue
                    else:
                        break
                
                #formatação de texto
                txt = f'''Peça {peca + 1}:
                '''
                for indice, codigo_peca in enumerate(produtos[peca]):
                    txt += f'''
    {indice + 1} - {codigo_peca}'''

                #verificação de tamanho da lista
                if len(produtos[peca]) == 3:
                    escolha = pg.confirm(f'{txt}', buttons=['Alterar código', 'Excluir código'])
                else:
                    escolha = pg.confirm(f'{txt}', buttons=['Alterar código', 'Adicionar código', 'Excluir código'])

                #escolhas
                if escolha == 'Alterar código':
                    # pegar codigo a ser corrigido
                    while True:
                        try:
                            code = int(pg.password(f'''Qual código você deseja corrigir? (Digite apenas o número):
                                                
    {txt}''', title='Cotação RPA', mask='')) - 1
                            codigo_corrigido = codigo_corrigido[code]
                        except ValueError:
                            pg.alert('Insira um valor numérico!')
                            continue
                        except IndexError:
                            pg.alert(f'Não há codigo {code + 1}!')
                            continue
                        else:
                            break

                    new_code = pg.password(f'{codigo_corrigido}', title='Cotação RPA', mask='')
                    if new_code:
                        produtos[peca][code] = new_code
                    continue
                
                elif escolha == 'Excluir código':

                    while True:
                        try:
                            codigo_excluido_index = int(pg.password(f'''Qual código você deseja excluir?
                                        
        {txt}''', mask='', title='Cotação RPA')) - 1
                            codigo_excluido = produtos[peca][codigo_excluido_index]
                        except IndexError:
                            pg.alert(f'Não existe código {codigo_excluido_index + 1}')
                            continue
                        except ValueError:
                            pg.alert(f'Digite um valor numérico!')
                            continue
                        else:
                            produtos.pop(codigo_excluido)
                            break

                elif escolha == 'Adicionar código':
                    new_code = pg.password('Digite o novo código.', title='Cotação RPA', mask='')
                    if new_code:
                        produtos[peca].append(new_code)
                    continue

            elif escolha == 'Adicionar mais peças':
                break
            
            elif escolha == 'Excluir Peça':
                #pegar peca a ser corrigida
                while True:
                    try:
                        peca = int(pg.password(f'''Peças cotadas:
    {txt}
    Qual peça você deseja excluir? (Digite apenas o número)''', title='Cotação RPA', mask='')) - 1
                        codigo_corrigido = produtos[peca]
                    except ValueError:
                        pg.alert('Insira um valor numérico!')
                        continue
                    except IndexError:
                        pg.alert(f'Não há peça {peca + 1}!')
                        continue
                    else:
                        break
                produtos.pop(peca)
            
            else:
                #return
                return produtos
                # print(produtos)
                # quit()

            continue
