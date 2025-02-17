import pandas as pd


# Função para normalizar o CNPJ (remover pontos, barra e traço)
def normalizar_cnpj(cnpj):
    if isinstance(cnpj, str):  # Verifica se o CNPJ é uma string
        return cnpj.replace('.', '').replace('/', '').replace('-', '').strip()
    return cnpj  # Retorna o valor original (ex.: NaN)


# Caminhos dos arquivos (ajuste de acordo com seus arquivos reais)
caminho_arquivo1 = r'C:\Users\gabriel.lima\Downloads\arquivo 1.xlsx'
caminho_arquivo2 = r'C:\Users\gabriel.lima\Downloads\arquivo 2.xlsx'

# Ler os arquivos com tratamento de erros
try:
    arquivo1 = pd.read_excel(caminho_arquivo1, engine='openpyxl')
    arquivo2 = pd.read_excel(caminho_arquivo2, engine='openpyxl')
except Exception as e:
    print(f"Erro ao carregar os arquivos: {e}")
    exit()

# Verificar se as colunas CNPJ e Venc existem nos arquivos
if 'CNPJ' not in arquivo1.columns or 'Venc' not in arquivo1.columns:
    print("Erro: As colunas 'CNPJ' e 'Venc' não foram encontradas no arquivo 1.")
    exit()
if 'CNPJ' not in arquivo2.columns or 'Venc' not in arquivo2.columns:
    print("Erro: As colunas 'CNPJ' e 'Venc' não foram encontradas no arquivo 2.")
    exit()

# Verificar se a coluna 'Empresa' existe no arquivo 1
if 'Empresa' not in arquivo1.columns:
    print("\nAviso: A coluna 'Empresa' não foi encontrada no arquivo 1. Usando valor padrão para Nome da Empresa.")
    # Criar uma coluna Empresa com valores padrão
    arquivo1['Empresa'] = "Empresa Desconhecida"

# Normalizar e garantir que o CNPJ esteja como string
arquivo1['CNPJ'] = arquivo1['CNPJ'].astype(str).apply(normalizar_cnpj)
arquivo2['CNPJ'] = arquivo2['CNPJ'].astype(str).apply(normalizar_cnpj)

# Remover valores ausentes em CNPJ
arquivo1.dropna(subset=['CNPJ'], inplace=True)
arquivo2.dropna(subset=['CNPJ'], inplace=True)

# Remover CNPJs inválidos (não numéricos)
arquivo1 = arquivo1[arquivo1['CNPJ'].str.isnumeric()]
arquivo2 = arquivo2[arquivo2['CNPJ'].str.isnumeric()]

# Renomear colunas 'Venc' para as datas corretas
arquivo1.rename(columns={'Venc': 'Data_Expiracao'}, inplace=True)
arquivo2.rename(columns={'Venc': 'Data_Vencimento'}, inplace=True)

# Converter as colunas de data para o formato datetime
arquivo1['Data_Expiracao'] = pd.to_datetime(arquivo1['Data_Expiracao'], errors='coerce').dt.normalize()
arquivo2['Data_Vencimento'] = pd.to_datetime(arquivo2['Data_Vencimento'], errors='coerce').dt.normalize()

# Debug: Certificar que as datas foram convertidas corretamente
print("\nDebug: Primeiras entradas do arquivo 1")
print(arquivo1[['CNPJ', 'Empresa', 'Data_Expiracao']].head())
print("\nDebug: Primeiras entradas do arquivo 2")
print(arquivo2[['CNPJ', 'Data_Vencimento']].head())

# Verificar interseção de CNPJs entre os dois DataFrames
common_cnpjs = set(arquivo1['CNPJ']).intersection(set(arquivo2['CNPJ']))
print(f"\nCNPJs em comum entre os dois arquivos: {len(common_cnpjs)}")

# Realizar merge somente nos registros com CNPJs em comum
merged_df = pd.merge(arquivo1, arquivo2, on='CNPJ', how='inner', suffixes=('_Arquivo1', '_Arquivo2'))

# Garantir que todas as colunas (Empresa, CNPJ, etc.) existam
print("\nDebug: Colunas após o merge:")
print(merged_df.columns)

# Comparar as datas de expiração
merged_df['Diferenca_Datas'] = merged_df['Data_Expiracao'] != merged_df['Data_Vencimento']

# Filtrar apenas registros com divergências de datas
diferencas_df = merged_df[merged_df['Diferenca_Datas']].copy()

# Garantir que a coluna "Empresa" esteja no DataFrame, mesmo que preenchida com valor padrão
if 'Empresa' not in diferencas_df.columns:
    diferencas_df['Empresa'] = "Empresa Desconhecida"

# Reduzir as colunas ao formato necessário
resultado_df = diferencas_df[['Empresa', 'CNPJ', 'Data_Vencimento']]

# Renomear colunas para o formato amigável e evitar SettingWithCopyWarning
resultado_df = resultado_df.rename(columns={'Empresa': 'Nome da Empresa', 'Data_Vencimento': 'Data'})

# Formatar a coluna de Data para o formato DD/MM/AAAA
resultado_df['Data'] = resultado_df['Data'].dt.strftime('%d/%m/%Y')

# Debug: Inspecionar registros finais para salvar
print("\nDebug: Registros finais a serem salvos no Excel:")
print(resultado_df)

# Verificar se o DataFrame final não está vazio antes de salvar
if not resultado_df.empty:
    try:
        resultado_df.to_excel('diferencas_datas.xlsx', index=False)
        print("\nArquivo 'diferencas_datas.xlsx' salvo com sucesso!")
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")
else:
    print("\nNenhuma diferença foi encontrada ou salva no arquivo.")
