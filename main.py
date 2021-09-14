import pyautogui
import time
import pyperclip
import pandas as pd
import datetime
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

print(20 * '=', 'Relatório de Vendas', 20 * '=')
while True:
    opc = int(input('1 - Resumo de Vendas\n2 - Relatório de Vendas via e-Mail\n'))
    if opc != 1 and opc != 2:
        print('Opção desconhecida.\nPor favor, digite uma opção válida!')
    else:
        break

if opc == 1:
    print()
    #Lembrar! A solicitação deve ser feita com try e catch.
    dataini = datetime.datetime.strptime(str(input('Digite a data inicial do período que gostaria de analisar'
                                                   '(dd/mm/aaaa):\n')), '%d/%m/%Y')
    datafim = datetime.datetime.strptime(str(input('Digite a data final do período que gostaria de analisar'
                                                   '(dd/mm/aaaa):\n')), '%d/%m/%Y')

    df = pd.read_excel(
        r'C:\Users\T460s\Documents\GitHub\projeto_python_envio_relatorios_automatizados/Vendas - Dez.xlsx')
    df_datas = pd.to_datetime(df['Data'], format='%b %d, %Y')

    apos_dataini = df['Data'] >= dataini
    antes_datafim = df['Data'] <= dataini
    df_formatado = apos_dataini & antes_datafim
    datas_filtradas = df.loc[df_formatado]


    print(20 * '-', 'Resumo Geral de Vendas', 20 * '-')
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

    print(10 * '-', 'Produto Mais Vendido', 10 * '-')
    print(f'-> Nome: {prod_mais_vendido_filtrado["Produto"][0]}')
    print(f'-> Quantidade: {prod_mais_vendido_filtrado["Quantidade"][0]}')
    print(f'-> Faturamento: {prod_mais_vendido_faturamento}')
    print()

    print(10 * '-', 'Loja Com Mais Vendas', 10 * '-')
    print(f'-> Nome: {loja_mais_vendas_filtrado["ID Loja"][0]}')
    print(f'-> Quantidade: {loja_mais_vendas_filtrado["Quantidade"][0]}')
    print(f'-> Faturamento: {loja_mais_vendas_faturamento}')
    print()

    print(17 * '-', 'TOTAL', 17 * '-')
    print(f'-> Produtos Vendidos: {qtd_prod}')
    print(f'-> Faturamento: {faturamento}')

'''pyautogui.PAUSE = 1

#abre navegador
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.alert("Vai começar, aperte OK e não mexa em nada")
pyautogui.hotkey('ctrl', 't')

#abre drive
link = "https://drive.google.com/drive/folders/1mhXZ3JPAnekXP_4vX7Z_sJj35VWqayaR/usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(15)'''

#produto = ''

'''for c, v in enumerate(maisvend):
    print(c)'''

#print(produtovalor)


'''pyautogui.press("winleft")
time.sleep(1)
pyautogui.write("chrome")
time.sleep(2)
pyautogui.press("enter")
time.sleep(5)
pyautogui.hotkey('ctrl', 't')
time.sleep(2)
pyautogui.write("mail.google.com")
pyautogui.press('enter')
time.sleep(5)
pyautogui.click(94, 227)'''
