# Projeto de Análise de Vendas
O projeto de análise de vendas em Python utiliza as bibliotecas Pandas e Openpyxl para realizar análises detalhadas sobre dados de vendas, incluindo faturamento por loja, quantidade de produtos vendidos e ticket médio. Além disso, incorpora a funcionalidade de envio automático de relatórios por e-mail usando a biblioteca win32com.client. Este script proporciona uma visão abrangente do desempenho de vendas, permitindo uma fácil compreensão e interpretação dos dados. Ideal para monitoramento e comunicação eficiente das métricas essenciais para os interessados no negócio.

## Pré-requisitos

Certifique-se de ter o Python e as bibliotecas necessárias instaladas. Você pode instalar as dependências executando:

```bash
pip install pandas openpyxl pywin32

## Como usar
Clone o repositório:
bash
Copy code
git clone https://github.com/seu-username/seu-repositorio.git
cd seu-repositorio

## Instale as dependências:
bash
Copy code
pip install -r requirements.txt

## Execute o script Python:
bash
Copy code
python seu_script.py

## Funcionalidades
Análise de Faturamento por Loja: Exibe o faturamento total por cada loja.
Quantidade de Produtos Vendidos por Loja: Mostra a quantidade total de produtos vendidos por loja.
Ticket Médio por Produto em Cada Loja: Calcula o ticket médio dos produtos em cada loja.
Envio de E-mail com Relatório: Utiliza o Outlook para enviar um e-mail com o relatório gerado.
