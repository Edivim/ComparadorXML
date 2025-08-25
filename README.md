Comparador de XMLs - Edivim

Este programa foi desenvolvido para comparar as chaves de acesso presentes em um arquivo Excel, 
com os nomes de arquivos XML em uma pasta. Ele identifica quais XMLs estão faltando com base, 
nas chaves listadas no Excel e gera um relatório em Excel com os resultados.

Funcionalidades

- Seleção de arquivo Excel (.xls ou .xlsx)
- Seleção de pasta contendo arquivos XML
- Comparação automática entre as chaves do Excel e os arquivos XML
- Exibição dos resultados em uma interface gráfica
- Geração de relatório `faltando_xmls.xlsx` com as chaves ausentes

Interface

O programa possui uma interface gráfica simples e intuitiva feita com `tkinter`, 
permitindo que qualquer usuário selecione os arquivos e execute a comparação com apenas alguns cliques.

Requisitos

- Python 3.10 ou superior
- Bibliotecas utilizadas:
  - `pandas`
  - `tkinter` (já incluída no Python padrão)
  - `openpyxl` (para salvar arquivos Excel)

Você pode instalar as dependências com:

pip install pandas openpyxl


Download do Instalador

Você pode baixar o programa completo (sem precisar instalar Python) usando o instalador com o nome abaixo:

ComparadorXML_Setup.exe

Após instalar, basta abrir o programa e começar a usar!

Como usar

1. Execute o script `comparador.py`
2. Clique em **Selecionar Excel** e escolha o arquivo com a coluna "Chave de Acesso"
3. Clique em **Selecionar Pasta** e escolha a pasta onde estão os arquivos XML
4. Clique em **Comparar**
5. O programa mostrará os resultados na tela e criará um arquivo chamado `faltando_xmls.xlsx` com as chaves que não foram encontradas

Os arquivos XML devem estar nomeados com a chave de acesso no início, como:
35190812345678901234550010000000011000000010-nfe.xml

O programa extrai a chave antes do primeiro hífen (-) para fazer a comparação.
