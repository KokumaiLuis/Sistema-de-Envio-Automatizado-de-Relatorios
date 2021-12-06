# Sistema de Envio Automatizado de Relatórios


## 💻 Funcionamento
A lógica analisa uma planilha de vendas em excel e, com base no período descrito pelo usuário, calcula:
1. Loja que mais vendeu e seu faturamento;
2. Produto mais vendido e seu faturamento;
3. Quantidade total de produtos vendidos;
4. Faturamento total.

Após isso envia um e-mail para o destinatário descrito pelo usuário com o relatório através do G-mail. (Também há a possibilidade de enviar e-mail com cópia para outros destinatários)

## :heavy_check_mark: Status do Projeto

:white_check_mark: Projeto Concluído! :white_check_mark:

## 📋 Pré-requisitos
Antes de começar, verifique se você atendeu aos seguintes requisitos:

* Versão mais recente de Python;
* Google Chrome ou Microsoft Edge (programa não foi testado em outros browsers);
* Conta no G-mail ativa e logada (existe a possibilidade de adaptar o programa para fazê-lo logar com base no login e senha que o usuário imputar, porém esta funcionalidade ainda não foi desenvolvida).


## 🚀 Instalação
Para ter um ambiente de desenvolvimento e testes do Sistema de Envio Automatizado de Relatórios, siga os passos a seguir:

```
1. Faça o download e a instalação da versão do Python mais recente em: (https://www.python.org/downloads/);
2. Faça a instalação dos módulos a seguir de acordo com sua IDE:
   1. Pandas (https://pandas.pydata.org/docs/getting_started/install.html);
   2. Pyautogui (https://pyautogui.readthedocs.io/en/latest/install.html);
   3. Time (https://pypi.org/project/python-time/);
   4. Pyperclip (https://pypi.org/project/pyperclip/);
   5. Datetime (https://pypi.org/project/DateTime/).
```


## ☕ Usando o Sistema de Envio Automatizado de Relatórios
Antes de testar o sistema, lembre-se de que você deve:

* Estar com conexão ativa à internet;
* Estar com sua conta ao G-mail já logada.

Para utilizar o Sistema de Envio Automatizado de Relatórios siga os seguintes passos:

1. Faça uma cópia deste repositório para o seu computador;
2. Abra o arquivo main.py em sua IDE;
3. Vá até a linha 25 do código e altere o caminho do diretório para o local em que está o seu arquivo excel "Vendas - Dez";
4. Execute o programa e siga as instruções da command window.

## 🗨️ Observações

* O código pode ser implementado para solicitar o site, login e senha em que o usuário deseja enviar o e-mail, qualquer dúvida abra uma issue para conversarmos sobre;
* O tempo de espera utilizado no código para executar as funções é um tempo médio padrão, ele pode ser alterado de acordo com a necessidade, qualquer dúvida abra uma issue para conversarmos sobre;
* O sistema foi testado nos navegadores Google Chrome e Microsoft Edge.

## 🛠️ Desenvolvido Com
```
PyCharm 2020.3.3 (Community Edition)
Build #PC-203.7148.72, built on January 27, 2021
Runtime version: 11.0.9.1+11-b1145.77 amd64
VM: OpenJDK 64-Bit Server VM by JetBrains s.r.o.
Windows 10 10.0
GC: ParNew, ConcurrentMarkSweep
Memory: 970M
Cores: 4
```

## 📫 Contribuindo com Sistema de Envio Automatizado de Relatórios
Para contribuir com Envio Automatizado de Relatórios, siga estas etapas:

1. Bifurque este repositório.
2. Crie um branch: `git checkout -b <nome_branch>`.
3. Faça suas alterações e confirme-as: `git commit -m '<mensagem_commit>'`
4. Envie para o branch original: `git push origin <nome_do_projeto> / <local>`
5. Crie a solicitação de pull.

Como alternativa, consulte a documentação do GitHub em [como criar uma solicitação pull](https://help.github.com/en/github/collaborating-with-issues-and-pull-requests/creating-a-pull-request).

## ✒️ Autores
* Luis Gustavo de Andrade Kokumai - (https://github.com/KokumaiLuis)

## 📝 Licença
Esse projeto está sob licença. Veja o arquivo [LICENÇA](LICENSE) para mais detalhes.

[⬆ Voltar ao topo](https://github.com/KokumaiLuis/Sistema-de-Envio-Automatizado-de-Relatorios)<br>
