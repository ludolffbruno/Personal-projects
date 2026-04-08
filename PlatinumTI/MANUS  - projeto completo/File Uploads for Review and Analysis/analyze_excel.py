import pandas as pd
import openpyxl
from openpyxl import load_workbook

# Analisar a estrutura da planilha
print("Analisando a estrutura da planilha CALCULO_ST_DIFAL_COMPLETA.xlsx...")

# Carregar a planilha com data_only=True para obter valores calculados
wb = load_workbook('/home/ubuntu/upload/CALCULO_ST_DIFAL_COMPLETA.xlsx', data_only=True)

# Listar todas as abas da planilha
print("\nAbas disponíveis na planilha:")
for sheet_name in wb.sheetnames:
    print(f"- {sheet_name}")

# Analisar a aba COMPLETA (mencionada no código Python)
if 'COMPLETA' in wb.sheetnames:
    sheet = wb['COMPLETA']
    
    # Obter os cabeçalhos (primeira linha)
    headers = []
    for cell in sheet[1]:
        headers.append(cell.value)
    
    print("\nCabeçalhos da aba COMPLETA:")
    for i, header in enumerate(headers):
        print(f"Coluna {chr(65+i)}: {header}")
    
    # Verificar o número de linhas e colunas
    max_row = sheet.max_row
    max_col = sheet.max_column
    print(f"\nNúmero de linhas: {max_row}")
    print(f"Número de colunas: {max_col}")
    
    # Analisar algumas linhas de exemplo
    print("\nExemplos de dados (primeiras 5 linhas):")
    for row in range(2, min(7, max_row + 1)):
        row_data = []
        for col in range(1, min(10, max_col + 1)):  # Limitando a 10 colunas para melhor visualização
            cell_value = sheet.cell(row=row, column=col).value
            row_data.append(str(cell_value))
        print(f"Linha {row}: {' | '.join(row_data)}")

# Carregar com pandas para análise adicional
try:
    # Tentar carregar a aba COMPLETA
    df = pd.read_excel('/home/ubuntu/upload/CALCULO_ST_DIFAL_COMPLETA.xlsx', sheet_name='COMPLETA')
    
    # Mostrar informações sobre os tipos de dados
    print("\nTipos de dados nas colunas:")
    print(df.dtypes)
    
    # Verificar valores únicos em colunas importantes
    if 'ST / DIFAL' in df.columns:
        print("\nValores únicos na coluna 'ST / DIFAL':")
        print(df['ST / DIFAL'].unique())
    
    # Verificar outras colunas mencionadas no código
    important_columns = ['Aquisição', 'Origem']
    for col in important_columns:
        if col in df.columns:
            print(f"\nValores únicos na coluna '{col}':")
            print(df[col].unique())
    
except Exception as e:
    print(f"Erro ao analisar com pandas: {e}")

# Analisar as colunas específicas mencionadas no código Python
print("\nAnálise das colunas específicas mencionadas no código Python:")
columns_to_check = {
    'C': 'ST / DIFAL',
    'D': 'Aquisição',
    'E': 'Origem',
    'F': 'ST Inclusa',
    'G': 'Origem Importada (4%)',
    'H': 'Origem Nacional (12%)',
    'I': 'Origem RJ (22%)'
}

if 'COMPLETA' in wb.sheetnames:
    sheet = wb['COMPLETA']
    for col_letter, col_name in columns_to_check.items():
        print(f"\nColuna {col_letter} ({col_name}):")
        values = set()
        for row in range(2, min(20, max_row + 1)):  # Analisar até 20 linhas
            cell_value = sheet[f"{col_letter}{row}"].value
            values.add(str(cell_value))
        print(f"Valores encontrados: {', '.join(values)}")
