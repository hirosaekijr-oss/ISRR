# ISRR
Sistema destinado a extrair dados das ISR Rede, que estão no formato PDF, armazenar esses dados em um BD e possibilitar a edição dos dados





Caminho da pasta onde devem ser colocados os arquivos e o caminho '\\\ELFS1\cos_programação\' deve ser mapeada como unidade Z:

<img width="1541" height="836" alt="ISRR" src="https://github.com/user-attachments/assets/e3a2e269-df19-42a7-8fc5-ae040ccb937f" />

# Pode ser necessário a instalação de algumas bibliotecas no Python 3.

# pip install pyodbc
# pip install openpyxl
# pip install pandas

# Para instalar a biblioteca "tika" é necessário instalar o "Java Development Kit (JDK)"

# sudo apt-get update
# sudo apt-get install default-jdk

# pip install tika


Caso o banco de dados seja movido para outro servidor, uma unidade de rede deve ser mapeada como 'Z:' e os arquivos abaixo devem ser colocados no diretório 'Z:\Access_COS\ISRR\'.

'1-EXECUTE_ESSE_PROGRAMA_PRIMEIRO.exe'      -compilar o programa  'PDF2XLSX.py' e salvar com o nome:  '1-EXECUTE_ESSE_PROGRAMA_PRIMEIRO.exe'

'2.exe'                                     -compilar o programa 'copia_output_to_padronizador.xlsm.py' e salvar com o nome:  '2.exe'

'3.exe'                                     -compilar o programa  'Executar Macro padronizador.py' e salvar com o nome '3.exe'

'4.exe'                                     -complilar o programa 'copiar_planilha.py' e salvar como '4.exe'

'5.exe'                                    -compilar o programa 'insere_dados.py' e salvar como '5.exe'

'ISRR.accdb' ou para melhor segurança alterar para 'ISRR.accdr'

'output.xls'

'padronizador.xlsm'

'padronizador.xlsx'

Devem ser colocados na pasta 'Z:\Access_COS\ISRR\'  e com a unidade Z: mapeada devidamente do servidor.


O banco de dados:

'ISRR_be.accdb'

Deve ser colocado na pasta 'Z:Access_COS\ISRR\ISRR_BD'

Deve ser criada a pasta 'PDF_LIDOS' dentro do diretório 'Z:\Access_COS\ISRR'


