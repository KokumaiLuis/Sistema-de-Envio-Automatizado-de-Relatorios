import pyautogui
import time
import pyperclip
import pandas as pd
import datetime
import locale

# Atribui local
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


# Verifica se o valor é uma data
def valida_data(texto):
    try:
        datetime.datetime.strptime(texto, '%d/%m/%Y')
        return True
    except ValueError:
        print('Formato incorreto, a data deve ser no formato "dd/mm/aaaa"')
        return False


# Lê excel
def le_excel():
    df_lido = pd.read_excel(
        r'C:\Users\T460s\Documents\GitHub\projeto_python_envio_relatorios_automatizados/Vendas - Dez.xlsx')
    return df_lido


# Retorna produto mais vendido de acordo com a data escolhida pelo usuário
def produto_mais_vendido(datas_filtradas):
    prod_mais_vendido = datas_filtradas[['Produto', 'Quantidade', 'Valor Final']].groupby('Produto').sum()
    prod_mais_vendido_filtrado = prod_mais_vendido.loc[
        prod_mais_vendido['Quantidade'] == prod_mais_vendido['Quantidade'].max()].reset_index()
    return prod_mais_vendido_filtrado


# Retorna faturamento do produto mais vendido de acordo com a data escolhida pelo usuário
def produto_mais_vendido_faturamento(prod_mais_vendido_filtrado):
    prod_mais_vendido_faturamento = locale.currency(prod_mais_vendido_filtrado['Valor Final'].sum())
    return prod_mais_vendido_faturamento


# Retorna faturamento da loja que mais vendeu de acordo com a data escolhida pelo usuário
def loja_mais_vendas(datas_filtradas):
    loja_mais_vendas = datas_filtradas[['ID Loja', 'Quantidade', 'Valor Final']].groupby('ID Loja').sum()
    loja_mais_vendas_filtrado = loja_mais_vendas.loc[
        loja_mais_vendas['Quantidade'] == loja_mais_vendas['Quantidade'].max()].reset_index()
    return loja_mais_vendas_filtrado


# Retorna faturamento da loja que mais vendeu de acordo com a data escolhida pelo usuário
def loja_mais_vendas_faturamento(loja_mais_vendas_filtrado):
    loja_mais_vendas_faturamento = locale.currency(loja_mais_vendas_filtrado['Valor Final'].sum())
    return loja_mais_vendas_faturamento


# Retorna o faturamento total de acordo com a data escolhida pelo usuário
def faturamento_total(datas_filtradas):
    faturamento = locale.currency(datas_filtradas['Valor Final'].sum())
    return faturamento


# Retorna a quantidade de produtos totais vendidos de acordo com a data escolhida pelo usuário
def qtd_total(datas_filtradas):
    qtd_prod = datas_filtradas['Quantidade'].sum()
    return qtd_prod


# Realiza a automação - abrir navegador
def automacao_abre_navegador():
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


# Realiza a automação - abrir gmail
def automacao_abre_gmail(destinatario):

    # Abre Gmail
    pyautogui.write("mail.google.com")
    pyautogui.press('enter')
    time.sleep(12)
    for c in range(0, 12):
        pyautogui.press('tab')
        time.sleep(1)
    pyautogui.press('enter')
    time.sleep(7)
    for c in destinatario:
        pyautogui.write(c)
        pyautogui.press('tab')


def automacao_escreve_assunto_email(assunto_email):
    pyautogui.press('tab')
    pyautogui.write(assunto_email)
    pyautogui.press('tab')


# Realiza a automação - Aplica emails em cópia
def automacao_email_em_copia(em_copia):
    pyautogui.hotkey('ctrl', 'shift', 'c')
    for c in em_copia:
        pyautogui.write(c)
        pyautogui.press('tab')


# Gera texto do corpo do e-mail
def corpo_email(prod_mais_vendido_filtrado,
                prod_mais_vendido_faturamento,
                loja_mais_vendas_filtrado,
                loja_mais_vendas_faturamento,
                qtd_prod,
                faturamento):
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
                  '\nhttps://github.com/KokumaiLuis/Sistema-de-Envio-Automatizado-de-Relatorios\n\n'
    return corpo_email


while True:
    # Título
    print(20 * '=', 'Relatório de Vendas', 20 * '=')
    # Faz while para verificação da resposta, laço para quando o usuário digita 1 ou 2.
    while True:
        opc = int(input('1 - Resumo de Vendas\n2 - Relatório de Vendas via e-Mail\n'))
        if opc != 1 and opc != 2:
            print('Opção desconhecida.\nPor favor, digite uma opção válida!')
        else:
            break
    if opc == 1:
        print()
        # Solicita datas para filtrar a tabela
        while True:
            data_inicial_input = input('Digite a data inicial do período que gostaria de analisar' '(dd/mm/aaaa):\n')
            if valida_data(data_inicial_input):
                break

        # Converte string em data
        dataini = datetime.datetime.strptime(data_inicial_input, '%d/%m/%Y')

        while True:
            data_final_input = input('Digite a data final do período que gostaria de analisar' '(dd/mm/aaaa):\n')
            if valida_data(data_final_input):
                break

        # Converte string em data
        datafim = datetime.datetime.strptime(data_final_input, '%d/%m/%Y')

        # Abre excel e guarda seus valores
        df = le_excel()

        # Filtra data de acordo com a data selecionada pelo usuário
        apos_dataini = df.loc[df['Data'] >= dataini]
        antes_datafim = apos_dataini.loc[apos_dataini['Data'] <= datafim]
        datas_filtradas = antes_datafim

        # Cabeçalho para mostrar os resultados
        print(20 * '-', 'Resumo Geral de Vendas', 20 * '-')
        # Mostra data filtrada
        print(f'Período: {dataini}', ' ~ ', f'{datafim}')
        print()

        # Mostra valores do produto mais vendido
        print(10 * '-', 'Produto Mais Vendido', 10 * '-')
        print(f'-> Nome: {produto_mais_vendido(datas_filtradas)["Produto"][0]}')
        print(f'-> Quantidade: {produto_mais_vendido(datas_filtradas)["Quantidade"][0]}')
        print(f'-> Faturamento: {produto_mais_vendido_faturamento(produto_mais_vendido(datas_filtradas))}')
        print()

        # Mostra valores da loja que mais vendeu
        print(10 * '-', 'Loja Com Mais Vendas', 10 * '-')
        print(f'-> Nome: {loja_mais_vendas(datas_filtradas)["ID Loja"][0]}')
        print(f'-> Quantidade: {loja_mais_vendas(datas_filtradas)["Quantidade"][0]}')
        print(f'-> Faturamento: {loja_mais_vendas_faturamento(loja_mais_vendas(datas_filtradas))}')
        print()

        # Mostra valores totais
        print(17 * '-', 'TOTAL', 17 * '-')
        print(f'-> Produtos Vendidos: {qtd_total(datas_filtradas)}')
        print(f'-> Faturamento: {faturamento_total(datas_filtradas)}')
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
            # Solicita datas para filtrar a tabela
            while True:
                data_inicial_input = input(
                    'Digite a data inicial do período que gostaria de analisar' '(dd/mm/aaaa):\n')
                if valida_data(data_inicial_input):
                    break

            # Converte string em data
            dataini = datetime.datetime.strptime(data_inicial_input, '%d/%m/%Y')

            while True:
                data_final_input = input('Digite a data final do período que gostaria de analisar' '(dd/mm/aaaa):\n')
                if valida_data(data_final_input):
                    break

            # Converte string em data
            datafim = datetime.datetime.strptime(data_final_input, '%d/%m/%Y')

            # Abre excel e guarda seus valores
            df = le_excel()

            # Filtra data de acordo com a data selecionada pelo usuário
            apos_dataini = df.loc[df['Data'] >= dataini]
            antes_datafim = apos_dataini.loc[apos_dataini['Data'] <= datafim]
            datas_filtradas = antes_datafim

            # Solicita navegador padrão para o usuário. Será este navegador que será utilizado para o envio do relatório
            navegador_padrao = str(input('Digite o nome do seu navegador de internet padrão:\n'))

            # Solicita o destinatário
            destinatario = []
            while True:
                temp_dest = str(input('Digite o(s) email(s) do(s) destinatário(s):("fim" p/ Parar)\n')).lower()
                if temp_dest not in 'fim':
                    destinatario.append(temp_dest)
                else:
                    break

            # Solicita o Assunto do E-mail
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
                    temp_em_copia = str(
                        input('Digite o(s) email(s) do(s) destinatário(s) em cópia:("fim" p/ Parar)\n')) \
                        .lower()
                    if temp_em_copia not in 'fim':
                        em_copia.append(temp_em_copia)
                    else:
                        break

                automacao_abre_navegador()
                automacao_abre_gmail(destinatario)
                automacao_email_em_copia(em_copia)
                automacao_escreve_assunto_email(assunto)
                pyperclip.copy(corpo_email(produto_mais_vendido(datas_filtradas),
                                           produto_mais_vendido_faturamento(produto_mais_vendido(datas_filtradas)),
                                           loja_mais_vendas(datas_filtradas),
                                           loja_mais_vendas_faturamento(loja_mais_vendas(datas_filtradas)),
                                           qtd_total(datas_filtradas),
                                           faturamento_total(datas_filtradas)))
                pyautogui.hotkey('ctrl', 'v')
                #pyautogui.hotkey('ctrl', 'enter')

            else:
                automacao_abre_navegador()
                automacao_abre_gmail(destinatario)
                automacao_escreve_assunto_email(assunto)
                pyperclip.copy(corpo_email(produto_mais_vendido(datas_filtradas),
                                           produto_mais_vendido_faturamento(produto_mais_vendido(datas_filtradas)),
                                           loja_mais_vendas(datas_filtradas),
                                           loja_mais_vendas_faturamento(loja_mais_vendas(datas_filtradas)),
                                           qtd_total(datas_filtradas),
                                           faturamento_total(datas_filtradas)))
                pyautogui.hotkey('ctrl', 'v')
                # pyautogui.hotkey('ctrl', 'enter')

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



