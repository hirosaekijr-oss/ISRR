from tika import parser
from openpyxl import Workbook
import os
import shutil
import subprocess
import win32com.client

usuario_atual = os.getenv('USERNAME')
caminho_downloads = os.path.join("C:\\Users", usuario_atual, "Downloads")
# caminho_arquivo_output = os.path.join(caminho_downloads, "output.xlsx")
caminho_arquivo_output = "Z:\Access_COS\ISRR\output.xlsx"



#print("O caminho da pasta Downloads do usuário", usuario_atual, "é:", caminho_downloads)

# Função para extrair texto de um arquivo PDF
def pdf_to_text(pdf_file):
    parsed_pdf = parser.from_file(pdf_file)
    text = parsed_pdf['content']
    return text

# Caminho da pasta contendo os arquivos PDF
folder_path = caminho_downloads

# Inicializa uma planilha do Excel
wb = Workbook()
ws = wb.active
ws.title = 'Planilha1'

# Pasta de destino para os arquivos lidos
# output_folder_path = os.path.join("C:\\Users", usuario_atual, "Downloads\PDF_LIDOS")
output_folder_path = "Z:\Access_COS\ISRR\PDF_LIDOS"
print("Aguarde a finalização da importação dos dados")
if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)

# Itera sobre os arquivos na pasta
for file_name in os.listdir(folder_path):
    if file_name.startswith('ISRR') and file_name.endswith('.pdf'):
        pdf_file_path = os.path.join(folder_path, file_name)
        text = pdf_to_text(pdf_file_path)
        
        # Verifica se o texto não é vazio
        if text:
            ws.append([file_name, text])
            
            # Move o arquivo para a pasta de arquivos lidos
            shutil.move(pdf_file_path, os.path.join(output_folder_path, file_name))
        else:
            ws.append([file_name, "Texto não pôde ser extraído"])

# Caminho de saída da planilha do Excel
output_excel_path = caminho_arquivo_output 

wb.save(output_excel_path)

#print("Conteúdo dos arquivos PDF que começam com ISRR salvos em", output_excel_path)
#print("Arquivos movidos para a pasta Z:\Access_COS\ISRR\PDF_LIDOS")

status = os.system('Z:\Access_COS\ISRR\\2.exe')
status = os.system('Z:\Access_COS\ISRR\\3.exe')
status = os.system('Z:\Access_COS\ISRR\\4.exe')
status = os.system('Z:\Access_COS\ISRR\\5.exe')
#print("Importação de dados finalizada com sucesso")
#final = input("ABRA O ISRR.ACCR E CLIQUE EM ATUALIZAR BD!!!!!")


