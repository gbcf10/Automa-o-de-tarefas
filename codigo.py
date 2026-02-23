import openpyxl
import pandas as pd

# Ler a planilha CSV usando pandas, especificando o separador como ponto e vírgula
tabela = pd.read_csv("energiaset.csv", delimiter=';')

# Exibir os nomes das colunas para verificar se 'leituras' existe
print("Colunas disponíveis no CSV:", tabela.columns)

# Nome da coluna que você deseja preencher
coluna_desejada = "leituras"  # Ajustado para remover o espaço em branco

# Verificar se a coluna desejada existe no DataFrame ( =coluna estrutura de dados)
if coluna_desejada not in tabela.columns:
    raise KeyError(f"A coluna '{coluna_desejada}' não foi encontrada no CSV.")

# Carregar o arquivo Excel onde você deseja preencher os dados
workbook = openpyxl.load_workbook('RMDOutt.xlsx')  # Substitua pelo nome do seu arquivo Excel
sheet = workbook.active

# Definir a coluna onde os dados serão preenchidos (por exemplo, a coluna F)
coluna_excel = 'F'  # Substitua pela letra da coluna desejada no Excel

# Preencher a coluna no Excel com os valores da coluna do CSV
# Começar na linha 6 assumindo que as linhas anteriores têm cabeçalho ou outros dados
for i, valor in enumerate(tabela[coluna_desejada], start=6):  
    cell = f"{coluna_excel}{i}"
    sheet[cell] = valor

# Salvar as alterações no arquivo Excel
workbook.save('RMDOutt.xlsx')  # Substitua pelo nome do seu arquivo Excel

print("Preenchimento da planilha concluído.")