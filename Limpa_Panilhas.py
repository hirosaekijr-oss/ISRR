import openpyxl

def limpar_valores(arquivo, planilha_nome):
    # Carregar a pasta de trabalho
    wb = openpyxl.load_workbook(arquivo, keep_vba=True)
    # Selecionar a planilha
    planilha = wb[planilha_nome]

    # Iterar sobre as células e limpar os valores (preservar fórmulas)
    for row in planilha.iter_rows():
        for cell in row:
            if not cell.has_style and not cell.data_type == 'f':  # Limpa apenas as células que não possuem estilo e não são fórmulas
                cell.value = None

    # Salvar as mudanças
    wb.save(arquivo)

# Caminhos dos arquivos
caminho_origem = r'Z:\Access_COS\ISRR\output.xlsx'
caminho_destino1 = r'Z:\Access_COS\ISRR\padronizador.xlsm'
caminho_destino2 = r'Z:\Access_COS\ISRR\padronizador.xlsx'

# Limpar valores nas planilhas específicas
limpar_valores(caminho_origem, 'Planilha1')
limpar_valores(caminho_destino1, 'Sheet1')
limpar_valores(caminho_destino2, 'Sheet2')
