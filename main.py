import pyautogui
import time
import pyperclip
import pandas as pd
import datetime
import locale

#Atribui local
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

while True:
    #Título
    print(20 * '=', 'Relatório de Vendas', 20 * '=')
    #Faz while para verificação da resposta, laço para quando o usuário digita 1 ou 2.
    while True:
        opc = int(input('1 - Resumo de Vendas\n2 - Relatório de Vendas via e-Mail\n'))
        if opc != 1 and opc != 2:
            print('Opção desconhecida.\nPor favor, digite uma opção válida!')
        else:
            break

    if opc == 1:
        print()
        #Lembrar! A solicitação deve ser feita com try e catch.
        #Solicita datas para filtrar a tabela
        dataini = datetime.datetime.strptime(str(input('Digite a data inicial do período que gostaria de analisar'
                                                       '(dd/mm/aaaa):\n')), '%d/%m/%Y')
        datafim = datetime.datetime.strptime(str(input('Digite a data final do período que gostaria de analisar'
                                                       '(dd/mm/aaaa):\n')), '%d/%m/%Y')

        #Lê excel
        df = pd.read_excel(
            r'C:\Users\T460s\Documents\GitHub\projeto_python_envio_relatorios_automatizados/Vendas - Dez.xlsx')
        #Converte os dados da coluna 'Data' para formate datetime
        df_datas = pd.to_datetime(df['Data'], format='%b %d, %Y')

        #Filtra data de acordo com a data selecionada pelo usuário
        apos_dataini = df['Data'] >= dataini
        antes_datafim = df['Data'] <= dataini
        df_formatado = apos_dataini & antes_datafim
        datas_filtradas = df.loc[df_formatado]

        #Cabeçalho para mostrar os resultados
        print(20 * '-', 'Resumo Geral de Vendas', 20 * '-')
        #Mostra data filtrada
        print(f'Período: {dataini}', ' ~ ', f'{datafim}')
        print()

        #Filtra produto mais vendido + quantidade + faturamento
        prod_mais_vendido = datas_filtradas[['Produto', 'Quantidade', 'Valor Final']].groupby('Produto').sum()
        prod_mais_vendido_filtrado = prod_mais_vendido.loc[
            prod_mais_vendido['Quantidade'] == prod_mais_vendido['Quantidade'].max()].reset_index()
        prod_mais_vendido_faturamento = locale.currency(prod_mais_vendido_filtrado['Valor Final'].sum())

        #Filtra loja que mais vendeu + quantidade + faturamento
        loja_mais_vendas = datas_filtradas[['ID Loja', 'Quantidade', 'Valor Final']].groupby('ID Loja').sum()
        loja_mais_vendas_filtrado = loja_mais_vendas.loc[
            loja_mais_vendas['Quantidade'] == loja_mais_vendas['Quantidade'].max()].reset_index()
        loja_mais_vendas_faturamento =  locale.currency(loja_mais_vendas_filtrado['Valor Final'].sum())

        #Faz a seleção do faturamento total + total de peças vendidas
        faturamento = locale.currency(datas_filtradas['Valor Final'].sum())
        qtd_prod = datas_filtradas['Quantidade'].sum()

        #Mostra valores do produto mais vendido
        print(10 * '-', 'Produto Mais Vendido', 10 * '-')
        print(f'-> Nome: {prod_mais_vendido_filtrado["Produto"][0]}')
        print(f'-> Quantidade: {prod_mais_vendido_filtrado["Quantidade"][0]}')
        print(f'-> Faturamento: {prod_mais_vendido_faturamento}')
        print()

        # Mostra valores da loja que mais vendeu
        print(10 * '-', 'Loja Com Mais Vendas', 10 * '-')
        print(f'-> Nome: {loja_mais_vendas_filtrado["ID Loja"][0]}')
        print(f'-> Quantidade: {loja_mais_vendas_filtrado["Quantidade"][0]}')
        print(f'-> Faturamento: {loja_mais_vendas_faturamento}')
        print()

        #Mostra valores totais
        print(17 * '-', 'TOTAL', 17 * '-')
        print(f'-> Produtos Vendidos: {qtd_prod}')
        print(f'-> Faturamento: {faturamento}')
        print()

    if opc == 2:
        while True:
            opc2 = str(input('Atenção!\n'
                             'A ação não poderá ser parada até o envio do relatório.\n'
                             'Verifique os seguintes pontos antes de iniciar:\n'
                             '1-Seu computador está com conexão à internet?\n'
                             '2-Sua conta do Gmail está logada?\n'
                             'Deseja realmente enviar relatório por email?(S/N)\n'))
            if opc2 not in 'SsNn':
                print('Opção inválida!\nPor favor, digite uma opção disponível.')
            else:
                break

        if opc2 in 'Ss':
            print()
            # Lembrar! A solicitação deve ser feita com try e catch.
            # Solicita datas para filtrar a tabela
            dataini = datetime.datetime.strptime(str(input('Digite a data inicial do período que gostaria de analisar'
                                                           '(dd/mm/aaaa):\n')), '%d/%m/%Y')
            datafim = datetime.datetime.strptime(str(input('Digite a data final do período que gostaria de analisar'
                                                           '(dd/mm/aaaa):\n')), '%d/%m/%Y')

            # Lê excel
            df = pd.read_excel(
                r'C:\Users\T460s\Documents\GitHub\projeto_python_envio_relatorios_automatizados/Vendas - Dez.xlsx')
            # Converte os dados da coluna 'Data' para formate datetime
            df_datas = pd.to_datetime(df['Data'], format='%b %d, %Y')

            # Filtra data de acordo com a data selecionada pelo usuário
            apos_dataini = df['Data'] >= dataini
            antes_datafim = df['Data'] <= dataini
            df_formatado = apos_dataini & antes_datafim
            datas_filtradas = df.loc[df_formatado]

            # Filtra produto mais vendido + quantidade + faturamento
            prod_mais_vendido = datas_filtradas[['Produto', 'Quantidade', 'Valor Final']].groupby('Produto').sum()
            prod_mais_vendido_filtrado = prod_mais_vendido.loc[
                prod_mais_vendido['Quantidade'] == prod_mais_vendido['Quantidade'].max()].reset_index()
            prod_mais_vendido_faturamento = locale.currency(prod_mais_vendido_filtrado['Valor Final'].sum())

            # Filtra loja que mais vendeu + quantidade + faturamento
            loja_mais_vendas = datas_filtradas[['ID Loja', 'Quantidade', 'Valor Final']].groupby('ID Loja').sum()
            loja_mais_vendas_filtrado = loja_mais_vendas.loc[
                loja_mais_vendas['Quantidade'] == loja_mais_vendas['Quantidade'].max()].reset_index()
            loja_mais_vendas_faturamento = locale.currency(loja_mais_vendas_filtrado['Valor Final'].sum())

            # Faz a seleção do faturamento total + total de peças vendidas
            faturamento = locale.currency(datas_filtradas['Valor Final'].sum())
            qtd_prod = datas_filtradas['Quantidade'].sum()

            #Solicita navegador padrão para o usuário. Será este navegador que será utilizado para o envio do relatório
            navegador_padrao = str(input('Digite o nome do seu navegador de internet padrão:\n'))

            #Solicita o destinatário
            destinatario = []
            while True:
                tempdest = str(input('Digite o(s) email(s) do(s) destinatário(s):("fim" p/ Parar)\n')).lower()
                if tempdest not in 'fim':
                    destinatario.append(tempdest)
                else:
                    break


            #Solicita o Assunto do E-mail
            assunto = str(input('Digite o assunto do email:\n'))

            while True:
                #Verifica a existência de emails em cópia
                opc3 = str(input('Terão emails em cópia?(S/N)\n'))
                if opc3 not in 'SsNn':
                    print('Opção inválida!\nPor favor, digite uma opção disponível.')
                else:
                    break
            if opc3 in 'Ss':
                em_copia = []
                while True:
                    tempemcopia = str(input('Digite o(s) email(s) do(s) destinatário(s) em cópia:("fim" p/ Parar)\n'))\
                        .lower()
                    if tempemcopia not in 'fim':
                        em_copia.append(tempemcopia)
                    else:
                        break

                pyautogui.PAUSE = 1

                pyautogui.alert("A automação iniciará em 3 segundos, aperte OK e após isso não mexa em nada.")
                time.sleep(3)

                # Abre Navegador
                pyautogui.press("winleft")
                time.sleep(3)
                pyautogui.write(navegador_padrao)
                time.sleep(3)
                pyautogui.press("enter")
                time.sleep(8)
                pyautogui.hotkey('ctrl', 't')
                time.sleep(3)

                # Abre Gmail
                pyautogui.write("mail.google.com")
                pyautogui.press('enter')
                time.sleep(12)
                pyautogui.click(94, 227)
                time.sleep(7)
                for c in destinatario:
                    pyautogui.write(c)
                    pyautogui.press('tab')
                pyautogui.hotkey('ctrl', 'shift', 'c')
                for c in em_copia:
                    pyautogui.write(c)
                    pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.write(assunto)
                pyautogui.press('tab')

                corpo_email = '----------Produto Mais Vendido----------\n' \
                              f'-> Nome: {prod_mais_vendido_filtrado["Produto"][0]}\n' \
                              f'-> Quantidade: {prod_mais_vendido_filtrado["Quantidade"][0]}\n' \
                              f'-> Faturamento: {prod_mais_vendido_faturamento}\n\n' \
                              \
                              '----------Loja Com Mais Vendas----------\n' \
                              f'-> Nome: {loja_mais_vendas_filtrado["ID Loja"][0]}\n' \
                              f'-> Quantidade: {loja_mais_vendas_filtrado["Quantidade"][0]}\n' \
                              f'-> Faturamento: {loja_mais_vendas_faturamento}\n\n' \
                              \
                              '-----------------TOTAL-----------------\n' \
                              f'-> Produtos Vendidos: {qtd_prod}\n' \
                              f'-> Faturamento: {faturamento}\n\n' \
                              \
                              'Este relatório foi gerado automaticamente por um programa em python. ' \
                              'Veja o código em:' \
                              '\nhttps://github.com/KokumaiLuis/Sistema-de-Envio-Automatizado-de-Relatorios\n\n' \

                pyperclip.copy(corpo_email)
                pyautogui.hotkey('ctrl', 'v')
                # pyautogui.hotkey('ctrl', 'enter')

            else:
                pyautogui.PAUSE = 1

                pyautogui.alert("A automação iniciará em 3 segundos, aperte OK e após isso não mexa em nada.")
                time.sleep(3)

                #Abre Navegador
                pyautogui.press("winleft")
                time.sleep(3)
                pyautogui.write(navegador_padrao)
                time.sleep(3)
                pyautogui.press("enter")
                time.sleep(8)
                pyautogui.hotkey('ctrl', 't')
                time.sleep(3)

                #Abre Gmail
                pyautogui.write("mail.google.com")
                pyautogui.press('enter')
                time.sleep(12)
                pyautogui.click(94, 227)
                time.sleep(7)
                for c in destinatario:
                    pyautogui.write(c)
                    pyautogui.press('space')
                pyautogui.press('tab')
                pyautogui.write(assunto)
                pyautogui.press('tab')

                corpo_email = '----------Produto Mais Vendido----------\n' \
                              f'-> Nome: {prod_mais_vendido_filtrado["Produto"][0]}\n' \
                              f'-> Quantidade: {prod_mais_vendido_filtrado["Quantidade"][0]}\n' \
                              f'-> Faturamento: {prod_mais_vendido_faturamento}\n\n' \
                              \
                              '----------Loja Com Mais Vendas----------\n' \
                              f'-> Nome: {loja_mais_vendas_filtrado["ID Loja"][0]}\n' \
                              f'-> Quantidade: {loja_mais_vendas_filtrado["Quantidade"][0]}\n' \
                              f'-> Faturamento: {loja_mais_vendas_faturamento}\n\n' \
                              \
                              '-----------------TOTAL-----------------\n' \
                              f'-> Produtos Vendidos: {qtd_prod}\n' \
                              f'-> Faturamento: {faturamento}\n\n' \
                              \
                              'Este relatório foi gerado automaticamente por um programa em python. ' \
                              'Veja o código em:' \
                              '\nhttps://github.com/KokumaiLuis/Sistema-de-Envio-Automatizado-de-Relatorios\n\n' \

                pyperclip.copy(corpo_email)
                pyautogui.hotkey('ctrl', 'v')
                #pyautogui.hotkey('ctrl', 'enter')

        else:
            print('Envio de relatório por E-mail cancelado.')
            time.sleep(1)

    while True:
        opc1 = str(input('Deseja iniciar novamente?(S/N)\n'))
        if opc1 not in 'SsNn':
            print('Opção inválida!\nPor favor, digite uma opção disponível.')
        else:
            break
    if opc1 in 'Nn':
        print('Encerrando Programa...')
        time.sleep(1)
        break
