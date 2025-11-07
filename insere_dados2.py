
import pyodbc
import pandas as pd
import os
from openpyxl import load_workbook

# Obtém o caminho absoluto do diretório onde o script está sendo executado
script_dir = os.path.dirname(os.path.abspath(__file__))
# file_path = os.path.join(script_dir, 'padronizador.xlsx')
file_path = "Z:\Access_COS\ISRR\padronizador.xlsx"

# Carregar os dados da planilha 'Sheet2' do arquivo Excel 'padronizador.xlsx' no mesmo diretório
df_excel = pd.read_excel(file_path, sheet_name='Sheet2')

# Estabelecer a conexão com o banco de dados no mesmo diretório
# db_file = os.path.join(script_dir, 'ISRR.accdb')
db_file = "Z:\Access_COS\ISRR\ISRR_BD\ISRR_be.accdb"


conn_str = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_file
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Inserir os dados do Excel no banco de dados, ignorando linhas com valores nulos
for index, row in df_excel.iterrows():
    if not row.isnull().values.any():
        cursor.execute("INSERT INTO tb_Apoio VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? ,?, ?, ?)", row['DOCUMENTO'], row['NUMERO'], row['ANO'], row['EMITENTE'], row['OBSERVACAO'], row['EQUIPAMENTO'], row['SERVIÇOS'], row['TELEFONE_SUPLENTE'], row['SUPLENTE'], row['TELEFONE_RESPONSAVEL'], row['RESPONSAVEL'], row['DATA_FINAL'], row['DATA_INICIO'], row['EQUIPAMENTO1'], row['LOCAL'])

conn.commit()
conn.close()

print("Dados do arquivo Excel inseridos no banco de dados, ignorando linhas com valores nulos.")





