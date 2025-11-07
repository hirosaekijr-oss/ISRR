import pandas as pd
import pyodbc

# Caminho do arquivo Excel
pasta = 'c:/teste'
arquivo_excel = 'padronizador.xlsx'
caminho_arquivo = f'{pasta}/{arquivo_excel}'

# Ler os dados da planilha Excel
df = pd.read_excel(caminho_arquivo, sheet_name='Sheet2')

# Conectar ao banco de dados Access
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:/teste/ISRR.accb;'
)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Verificar duplicatas na tabela no banco de dados
    cursor.execute('SELECT * FROM tb_ISRR')
    existing_data = cursor.fetchall()

    # Inserir os dados se não houver duplicatas
    duplicates = [row for row in df.iterrows() if row in existing_data]
    if not duplicates:
        for index, row in df.iterrows():
            cursor.execute('INSERT INTO tb_ISRR (coluna1, coluna2) VALUES (?, ?)', row['coluna1'], row['coluna2'])
        conn.commit()
        print('Dados inseridos com sucesso!')
    else:
        print('Dados duplicados encontrados. Não foi possível inserir os dados.')
except pyodbc.Error as ex:
    print('Erro de conexão:', ex)
