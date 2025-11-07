import pyodbc
import pandas as pd
import os
from openpyxl import load_workbook
contador = 0

script_dir = os.path.dirname(os.path.abspath(__file__))
# file_path = os.path.join(script_dir, 'padronizador.xlsx')
file_path = "Z:\Access_COS\ISRR\padronizador.xlsx"
print(file_path)
print(script_dir)


df_excel = pd.read_excel(file_path, sheet_name='Sheet2')


# db_file = os.path.join(script_dir, 'ISRR_be.accdb')

db_file = "Z:\Access_COS\ISRR\ISRR_BD\ISRR_be.accdb"
conn_str = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_file
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
print(db_file)

for index, row in df_excel.iterrows():
    if not row.isnull().values.any():
        cursor.execute("INSERT INTO tb_Apoio VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? ,?, ?, ?)", row['DOCUMENTO'], row['NUMERO'], row['ANO'], row['EMITENTE'], row['OBSERVACAO'], row['EQUIPAMENTO'], row['SERVIÇOS'], row['TELEFONE_SUPLENTE'], row['SUPLENTE'], row['TELEFONE_RESPONSAVEL'], row['RESPONSAVEL'], row['DATA_FINAL'], row['DATA_INICIO'], row['EQUIPAMENTO1'], row['LOCAL'])
        contador = contador + 1
conn.commit()
conn.close()

print("Dados do arquivo Excel inseridos no banco de dados, ignorando linhas com valores nulos.")
print(contador)

input("Aperte 'Enter' para fechar")



