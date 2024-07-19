# Leitor de Contratos
![](/screenshots/screenshot1.png)

**Um código simples, que aceita alguns contratos como entrada e gera um relatório com as informações desejadas.**

**Consiste na verdade de um ambiente virtual (venv) com algumas poucas libs instaladas, apenas para atender esse propósito.**

**Obs: Esse código foi feito originalmente para alimentar um robô feito em UiPath para o cadastro automático de clientes contidos nos contratos.**

## Para que o código servisse bem seu propósito alguns passos foram tomados

**- 'import' das bibliotecas que eram necessárias para o projeto:**
![](/screenshots/screenshot2.png)

- As bibliotecas 'PyPDF2' e 'pandas' tiveram que ser instaladas via 'pip install' com o venv ativado, assim como a biblioteca 'openpyxl' para manuseio do arquivo excel.


**- Definição da DataFrame que iria conter o relatório:**
![](/screenshots/screenshot3.png)

- Aqui foram definidas cada uma das colunas que iriam compor a Df. Esse exemplo contempla os dados que foram buscados nesses contratos, mas poderiam ser outros.


**- REGEX para buscar os dados nos contratos:**
![](/screenshots/screenshot4.png)

- Para cada arquivo encontrado na pasta '.\Contratos' é feita a busca por meio de expressões regulares, no fim da iteração esses dados são adicionados como uma linha na DataFrame.


**- Geração do relatório de contratos:**
![](/screenshots/screenshot5.png)
![](/screenshots/screenshot6.png)

- No fim é definido um nome e caminho para o relatório e o mesmo é gerado pelo **pandas** + **openpyxl**
