import pyodbc
import os

# Obtém o caminho do diretório onde o script Python está sendo executado
script_dir = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.join(script_dir, 'ISRR_be.accdb')

# Estabelece a conexão com o banco de dados
conn_str = r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path
conn = pyodbc.connect(conn_str)

cursor = conn.cursor()

# Verifica as tabelas existentes no banco de dados
tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
print("Tabelas disponíveis no banco de dados:", tables)

# Se 'tb_ISRR' estiver na lista de tabelas, executa a consulta
if 'tb_ISRR' in tables:
    cursor.execute('SELECT * FROM tb_ISRR')

    for row in cursor.fetchall():
        print(row)
else:
    print("A tabela 'tb_ISRR' não foi encontrada no banco de dados.")

conn.close()
