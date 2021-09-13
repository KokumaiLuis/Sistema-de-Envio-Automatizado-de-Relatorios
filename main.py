import pyautogui
import time
import pyperclip
import pandas as pd
import datetime

df = pd.read_excel(
        r'C:\Users\T460s\Documents\GitHub\projeto_python_envio_relatorios_automatizados/Vendas - Dez.xlsx')
print(pd.to_datetime(df['Data'], format='%b %d, %Y'))

print(20 * '=', 'Relatório de Vendas', 20 * '=')
while True:
    opc = int(input('1 - Resumo de Vendas\n2 - Relatório de Vendas via e-Mail\n'))
    if opc != 1 and opc != 2:
        print('Opção desconhecida.\nPor favor, digite uma opção válida!')
    else:
        break

if opc == 1:
    print()
    dataini = datetime.datetime.strptime(str(input('Digite a data inicial do período que gostaria de analisar'
                                                   '(dd/mm/aaaa):\n')), '%d/%m/%Y')
    datafim = datetime.datetime.strptime(str(input('Digite a data final do período que gostaria de analisar'
                                                   '(dd/mm/aaaa):\n')), '%d/%m/%Y')

    df = pd.read_excel(
        r'C:\Users\T460s\Documents\GitHub\projeto_python_envio_relatorios_automatizados/Vendas - Dez.xlsx')
    df_datas = pd.to_datetime(df['Data'], format='%b %d, %Y')
    df_formatado = df.loc[dataini <= df_datas <= datafim]

    print(df_formatado)

    #df_formatado = df[['Código Venda', 'Data', 'ID Loja', 'Produto', 'Quantidade', 'Valor Unitário', 'Valor Final']]

    print(10 * '-', 'Resumo Geral de Vendas', 10 * '-')


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

df = pd.read_excel(r'C:\Users\T460s\Documents\GitHub\projeto_python_envio_relatorios_automatizados/Vendas - Dez.xlsx')
produtovalor = df[['Produto', 'Valor Final']]
#print(df)

faturamento = df['Valor Final'].sum()
qtd_prod = df['Quantidade'].sum()

maisvend = df[['Produto', 'Quantidade']].groupby('Produto').sum()
maisvend1 = maisvend.loc[maisvend['Quantidade'] == maisvend['Quantidade'].max()].reset_index()


#produto = ''

'''for c, v in enumerate(maisvend):
    print(c)'''

print('Produto mais vendido:')
print(maisvend1['Quantidade'][0])
print(maisvend1['Produto'][0])

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
