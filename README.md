# Automação de Preenchimento de Formulário - Lists SharePoint

Este projeto foi desenvolvido para automatizar o processo de preenchimento de formulários no SharePoint a partir de dados contidos em uma planilha Excel. Utilizando o Selenium para controlar o navegador e o Tkinter para criar uma interface gráfica amigável, o sistema facilita o envio repetido de dados para o SharePoint, reduzindo erros manuais e economizando tempo.

FUNCIONALIDADES PRINCIPAIS

1. Automação de Preenchimento de Formulário
O sistema extrai dados de uma planilha Excel e preenche automaticamente um formulário hospedado no SharePoint. Cada campo do formulário é mapeado para uma coluna específica da planilha.

2. Interface Gráfica com Tkinter
O usuário pode iniciar o processo de automação através de uma interface gráfica simples. A interface permite abrir o navegador, acessar o formulário e iniciar o processo de preenchimento com apenas alguns cliques.

3. Integração com Excel
Os dados são extraídos diretamente de um arquivo Excel, tornando o processo flexível e fácil de atualizar. O sistema suporta planilhas estruturadas com 17 colunas, conforme o modelo fornecido.

4. Suporte a Execução Paralela
O preenchimento dos dados no formulário ocorre em uma thread separada, garantindo que a interface gráfica continue responsiva enquanto o processo de automação está em andamento.

TECNOLOGIAS UTILIZADAS

1. Python: Linguagem principal utilizada no projeto.

2. Selenium: Utilizado para controlar o navegador e preencher automaticamente os campos do formulário no SharePoint.

3. Tkinter: Biblioteca para criar a interface gráfica do usuário.

4. Pandas: Utilizado para manipulação e leitura dos dados da planilha Excel.

5. Excel: Fonte de dados para o preenchimento do formulário.

COMO FUNCIONA

Leitura da Planilha Excel: O sistema carrega uma planilha Excel que contém os dados necessários para preencher o formulário. A planilha deve ter as seguintes colunas:
DATA PROVISÃO, CLIENTE, UND NEGÓCIO, I.C, TIPO DOC, NÚM DOC, CLASSIFICAÇÃO, RECEITA BRUTA, ICMS, ISS, PIS, COFINS, CPRB, RECEITA LÍQUIDA, OBSERVAÇÃO, STATUS, LINK

Abertura do Navegador: O sistema abre automaticamente o navegador (Chrome) e acessa a página do SharePoint onde o formulário está hospedado.

Preenchimento Automático: Utilizando os dados da planilha, o sistema preenche cada campo do formulário e clica no botão "Enviar". Após o envio, o sistema aguarda o carregamento do formulário novamente para processar a próxima linha de dados.

Interface Gráfica: A interface permite ao usuário abrir o navegador e iniciar o processo de preenchimento de forma simples e rápida, tudo com alguns cliques.

CONCLUSÃO

Este projeto oferece uma solução prática e eficiente para automatizar o preenchimento de formulários no SharePoint, utilizando dados provenientes de uma planilha Excel. Com uma interface gráfica intuitiva e integração com Selenium, o sistema simplifica o processo de inserção de grandes volumes de dados, tornando-o mais rápido e menos propenso a erros manuais.
