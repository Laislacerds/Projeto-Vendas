# Projeto de Análise de Vendas

Este é um projeto de análise de vendas usando Python, Pandas e Openpyxl. Ele realiza algumas análises sobre os dados de vendas e envia um relatório por e-mail utilizando a biblioteca win32com.client.

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
