# Instalar a biblioteca necessária
!pip install openpyxl

import pandas as pd
from google.colab import files

# Fazer o upload do arquivo
uploaded = files.upload()

# Carregar as duas sheets
produtos = pd.read_excel('EstoquexFaturamento.xlsx', sheet_name='Produtos')
faturamento = pd.read_excel('EstoquexFaturamento.xlsx', sheet_name='Faturamento 2024')

# Remover espaços em branco nos nomes das colunas
produtos.columns = produtos.columns.str.strip()
faturamento.columns = faturamento.columns.str.strip()

# 1. Remover espaços em branco nas colunas de texto
produtos['Produto'] = produtos['Produto'].str.strip()
produtos['Serviço'] = produtos['Serviço'].str.strip()

# 2. Remover símbolos de moeda (R$) e converter para float
produtos['Preço'] = produtos['Preço'].replace({'R\$ ': '', '': '0'}, regex=True).astype(float)
produtos['Valor Pago'] = produtos['Valor Pago'].replace({'R\$ ': '', '': '0'}, regex=True).astype(float)

# 3. Verificar se a coluna de validade está numérica
produtos['Validade (meses)'] = pd.to_numeric(produtos['Validade (meses)'], errors='coerce')

# Exibir os dados tratados da planilha Produtos
print(produtos.head())

# 4. Converter a coluna "Data" para o formato datetime
faturamento['Data'] = pd.to_datetime(faturamento['Data'], errors='coerce')

# Criar um ExcelWriter para salvar as sheets em um novo arquivo
with pd.ExcelWriter('nova_planilha_tratadat.xlsx') as writer:
    produtos.to_excel(writer, sheet_name='Produtos', index=False)  # Salvar a sheet 'Produtos'
    faturamento.to_excel(writer, sheet_name='Faturamento 2024', index=False)  # Salvar a sheet 'Faturamento 2024'


# Baixar a nova planilha
files.download('nova_planilha_tratadat.xlsx')
