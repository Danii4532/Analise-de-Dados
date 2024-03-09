import os
import time
import tkinter
from tkinter import filedialog
import pandas as pd

def Tentar_novamente():
    op = input('Escolha \n1 - Continuar \n2 - Sair\n:')
    if op == '1':
        Principal()
    elif op == '2':
        exit()
    else:
        print('Opcão inválida')
        Tentar_novamente()

def salvar_planilha(arquivo):
    extensao = '.xlsx'
    salvar = input('Deseja salvar o arquivo em uma planilha do Excel nova?\n1 - Sim\n2 - Não\n')
    if salvar == '1':
        nome = input('Digite o nome da nova planilha: ') + extensao
        caminho = filedialog.askdirectory()

        # Remova o índice nomeado, se houver
        arquivo.reset_index(drop=True, inplace=True)

        arquivo.to_excel(os.path.join(caminho, nome), index=False)
        Tentar_novamente()
    elif salvar == '2':
        Tentar_novamente()


def Principal():
    print('Selecione a planilha que será usada')
    time.sleep(2)
    janela = tkinter.Tk()
    janela.title('Selecione a planilha')
    janela.withdraw()
    arquivo = filedialog.askopenfilename()
    caminho = os.path.join(arquivo)
    if arquivo:
        pass
    else:
        print('Erro ao abrir arquivo')
        Tentar_novamente()
    if arquivo.endswith('.xlsx'):
        pass
    else:
        print('Só podem ser usados arquivos com extensão ".xlsx"')
        Tentar_novamente()
    planilha = pd.read_excel(caminho)
    colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

    # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
    colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

    # Se houver colunas vazias, imprima o nome das colunas
    if not colunas_vazias.empty:
        for coluna in colunas_vazias:
            if coluna in planilha.columns:
                planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame
    print('---------------- DS Analyzer ----------------')
    visualizar = input('Deseja visualizar a planilha: \n1 - Sim\n2 - Não\n:')
    if visualizar == '1':


        print(planilha)

    elif visualizar == '2':
        pass

    else:
        print('Opção inválida')
        Tentar_novamente()

    for colunas in planilha:
        planilha[colunas] = planilha[colunas].astype('object')
    print('''
    Escolha:
    1 - Selecionar Maior Valor e Menor Valor
    2 - Selecionar Media
    3 - Quantidade linhas
    4 - Alterar Valor Celula
    5 - Alterar Valor Coluna
    6 - Adicionar Coluna com valores
    7 - Apagar Coluna
    8 - Apagar Célula
    9 - Somar valores da coluna
    
    ''')
    op = input(':')

    match op:
        case '1':
            # 1 - Selecionar Maior Valor
            #solicitando nome da coluna que sera usada
            print('Use essa função em colunas que possuam apenas números')
            coluna = input('Digite o nome exato da coluna que sera usada: ')
            #verificando se coluna existe na planilha,se sim, codigo continua, se nao, encerrar codigo
            if coluna not in planilha: print('Essa coluna não existe na planilha') ; Tentar_novamente()

            planilha[coluna] = pd.to_numeric(planilha[coluna], errors='coerce')
            valores_numericos = planilha[coluna].dropna()

            # Verifica se existem valores numéricos na coluna
            if valores_numericos.empty:
                print('Não há valores numéricos nesta coluna')
                Tentar_novamente()
            else:
                pass
            maior = planilha[coluna].max()
            menor = planilha[coluna].min()
            print(f'Coluna {coluna}')
            print(f'Menor valor:{menor}')
            print(f'Maior valor: {maior}\n')
            Tentar_novamente()

        case '2':
            # 2 - Selecionar Media
            print('Use essa função em colunas que possuam apenas números')
            coluna = input('Digite o nome exato da coluna que sera usada: ')
            # verificando se coluna existe na planilha,se sim, codigo continua, se nao, encerrar codigo
            if coluna not in planilha: print('Essa coluna não existe na planilha') ; Tentar_novamente()
            planilha[coluna] = pd.to_numeric(planilha[coluna], errors='coerce')
            valores_numericos = planilha[coluna].dropna()

            # Verifica se existem valores numéricos na coluna
            if valores_numericos.empty:
                print('Não há valores numéricos nesta coluna')
                Tentar_novamente()
            else:
                pass
            media = planilha[coluna].median()
            print(f'Coluna {coluna}')
            print(f'A média é: {media}\n')
            Tentar_novamente()
        case '3':
            # 3 - Quantidade linhas
            coluna = input('Digite o nome exato da coluna que sera usada: ')
            # verificando se coluna existe na planilha,se sim, codigo continua, se nao, encerrar codigo
            if coluna not in planilha: print('Essa coluna não existe na planilha') ; Tentar_novamente()
            quantidade = planilha[coluna].count()
            print(f'Tirando "{coluna}" são: {quantidade} linhas')
            Tentar_novamente()
        case '4':
            # 4 - Alterar Valor Celula
            coluna = input('Digite o nome exato da coluna que sera usada: ')
            if coluna not in planilha: print('Essa coluna não existe na planilha'); Tentar_novamente()
            celula = int(input('Digite o numero da celula sera alterada: '))
            # verificando se coluna existe na planilha,se sim, codigo continua, se nao, encerrar codigo
            if celula not in planilha[coluna]: print(f'Essa celula não existe na coluna {coluna}'); Tentar_novamente()

            valor_antigo = planilha[coluna][celula]
            novo_valor = input('Digite o novo valor que a celula terá: ')

            colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

            # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
            colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

            # Se houver colunas vazias, imprima o nome das colunas
            if not colunas_vazias.empty:
                for coluna in colunas_vazias:
                    if coluna in planilha.columns:
                        planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame
            # Verificando se usuário deseja visualizar a planilha modificada
            visualizar = input('Deseja visualizar a planilha modificada: \n1 - Sim \n2 - Não')

            if visualizar == '1':
                print(planilha)
                salvar_planilha(planilha)
            elif visualizar == '2':
                salvar_planilha(planilha)
                Tentar_novamente()

            else:
                print('Opção inválida')
                Tentar_novamente()
            novo1 = planilha.loc[celula,coluna] = novo_valor
            planilha[coluna][celula] == novo1
            salvar_planilha(planilha)
            print(f'coluna {coluna}')
            print(f'Antigo valor: {valor_antigo} \nNovo Valor: {novo1}')

            Tentar_novamente()

        case '5':
            #5 - Alterar Valor Coluna
            print('Essa função irá alterar todas as linhas da coluna')
            coluna = input('Digite o nome exato da coluna que sera alterada: ')
            if coluna not in planilha: print('Essa coluna não existe na planilha'); Tentar_novamente()
            print(f'Essa coluna possui {planilha[coluna].count()} linhas')
            quantidade = len(planilha[coluna])
            novo_valor = input('Digite o novo valor: ')
            for i,cada in enumerate(planilha[coluna]):
                planilha.loc[i,coluna] = novo_valor

            #verificar se a planilha possui coluna vazia e deleta
            colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

            # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
            colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

            # Se houver colunas vazias, imprima o nome das colunas
            if not colunas_vazias.empty:
                for coluna in colunas_vazias:
                    if coluna in planilha.columns:
                        planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame
            #Verificando se usuário deseja visualizar a planilha modificada
            visualizar = input('Deseja visualizar a planilha modificada: \n1 - Sim \n2 - Não')

            if visualizar == '1':
                print(planilha)
                salvar_planilha(planilha)
            elif visualizar == '2':
                salvar_planilha(planilha)
                Tentar_novamente()

            else:
                print('Opção inválida')
                Tentar_novamente()



        case '6':
            #6 - Adicionar Coluna com valores
            coluna = input('Digite o nome da coluna que sera criada: ')
            if coluna in planilha: print('Essa coluna ja existe na planilha') ; Tentar_novamente()
            op = input('Escolha \n1 - Mesmo valor para todas as linhas \n2 - Valor diferente para cada linha\n:')
            linhas = input('Quantas linhas a coluna tera: ')
            if linhas.isdigit() == False:
                print('Insira apenas números de 0 a 9')
                Tentar_novamente()
            else:
                pass
            if op == '1':
                novo_valor = input('Digite o novo valor: ')
                for i,cada in enumerate(range(int(linhas))):
                    planilha.loc[i,coluna] = novo_valor
                visualizar = input('Deseja visualizar o resultado \n1 - Sim \n2 - Não')
                if visualizar == '1': print(f'Coluna {coluna} modificada') ; print(planilha[coluna])
                elif visualizar == '2': Tentar_novamente()
                else: print('Opção Inválida')

                colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

                # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
                colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

                # Se houver colunas vazias, imprima o nome das colunas
                if not colunas_vazias.empty:
                    for coluna in colunas_vazias:
                        if coluna in planilha.columns:
                            planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame
                else:
                    pass

                salvar_planilha(planilha)

            elif op == '2':
                for cada in range(int(linhas)):
                    valor = input(f'Digite o {int(cada+1)}º valor: ')
                    planilha.loc[cada,coluna] = valor

                visualizar = input('Deseja visualizar o resultado \n1 - Sim \n2 - Não')
                if visualizar == '1':
                    print(f'Coluna {coluna} modificada')
                    print(planilha[coluna])
                    colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

                    # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
                    colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

                    # Se houver colunas vazias, imprima o nome das colunas
                    if not colunas_vazias.empty:
                        for coluna in colunas_vazias:
                            if coluna in planilha.columns:
                                planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame
                    else:
                        pass

                    salvar_planilha(planilha)

                elif visualizar == '2':
                    colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

                    # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
                    colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

                    # Se houver colunas vazias, imprima o nome das colunas
                    if not colunas_vazias.empty:
                        print("As seguintes colunas têm valores vazios ou estão totalmente vazias, exceto A1:")
                        for coluna in colunas_vazias:
                            if coluna in planilha.columns:
                                planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame
                        print('Colunas vazias excluídas com sucesso.')
                    else:
                        print("Não há colunas com valores vazios ou totalmente vazias, exceto A1.")

                    salvar_planilha(planilha)


                    Tentar_novamente()
                else:
                    print('Opção Inválida')

            else:
                print('Opção Invalida')

        case '7':
            #7 - Apagar Coluna
            coluna = input('Digite o nome da coluna: ')
            if coluna not in planilha: print('Essa coluna não existe na planilha'); Tentar_novamente()
            planilha = planilha.drop(coluna,axis=1)

            #verificando se coluna foi excluida
            if all(planilha):
                print(f'{coluna} Excluido com sucesso')
            else:
                print(f'Erro ao excluir coluna {coluna}')

            visualizar = input('Deseja visualizar o resultado \n1 - Sim \n2 - Não')
            if visualizar == '1':
                colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

                # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
                colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

                # Se houver colunas vazias, imprima o nome das colunas
                if not colunas_vazias.empty:
                    for coluna in colunas_vazias:
                        if coluna in planilha.columns:
                            planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame
                else:
                    pass
                print(planilha[:])
                salvar_planilha(planilha)

            elif visualizar == '2':
                colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

                # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
                colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

                # Se houver colunas vazias, imprima o nome das colunas
                if not colunas_vazias.empty:
                    for coluna in colunas_vazias:
                        if coluna in planilha.columns:
                            planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame

                else:
                    pass

                salvar_planilha(planilha)
            else:
                print('Opção Inválida')
        case '8':
            # 8 - Apagar Celula
            coluna = input('Digite o nome da coluna: ')
            if coluna not in planilha:
                print('Essa coluna não está presente na planilha')
                Tentar_novamente()
            print(planilha[coluna])
            celula = int(input('Digite a posição da celula: '))
            if celula not in planilha[coluna]:
                print('Essa célula não está presente na coluna especificada')
                Tentar_novamente()
            planilha.index = range(1, len(planilha[coluna]) + 1)
            for cada in range(1,len(planilha[coluna])):
                if cada == celula:
                    planilha.loc[cada,coluna] = None

            novo = planilha.loc[celula,coluna] = None
            print(f' removida com sucesso')
            print(planilha[coluna])

            colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

            # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
            colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

            # Se houver colunas vazias, imprima o nome das colunas
            if not colunas_vazias.empty:
                for coluna in colunas_vazias:
                    if coluna in planilha.columns:
                        planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame

            else:
                pass

            salvar_planilha(planilha)



        case '9':
            #9 - somar valores da coluna
            print('Use essa função em colunas que possuam apenas números')
            coluna = input('Digite o nome da coluna: ')
            if coluna not in planilha: print('Essa coluna não existe na planilha'); Tentar_novamente()

            planilha[coluna] = pd.to_numeric(planilha[coluna], errors='coerce')
            valores_numericos = planilha[coluna].dropna()

            colunas_nao_vazias = planilha.iloc[:, 1:].notnull().any()

            # Selecione colunas com valores nulos, mas apenas aquelas que também têm pelo menos um valor não nulo
            colunas_vazias = planilha.columns[planilha.columns != planilha.columns[0]][~colunas_nao_vazias]

            # Se houver colunas vazias, imprima o nome das colunas
            if not colunas_vazias.empty:
                for coluna in colunas_vazias:
                    if coluna in planilha.columns:
                        planilha = planilha.drop(coluna, axis=1)  # Exclui a coluna vazia do DataFrame

            else:
                pass

            # Verifica se existem valores numéricos na coluna
            if valores_numericos.empty:
                print('Não há valores numéricos nesta coluna')
                visualizar = input('Deseja visualizar a planilha: \n1 - Sim \n2 - Não')
                if visualizar == '1':
                    print(planilha)
                    Tentar_novamente()
                elif visualizar == '2':
                    Tentar_novamente()
                else:
                    print('Opção inválida')
                    Tentar_novamente()
            else:
                pass
            # Some apenas os valores numéricos da coluna
            soma_numerica = planilha[coluna].sum(skipna=True)

            print(f"Soma dos valores numéricos na coluna {coluna} é:", soma_numerica)
            Tentar_novamente()

        case _:
            print('Opção inválida')
            Tentar_novamente()
Principal()