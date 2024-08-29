import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def compare_excel_sheets(file1, file2, output_file):
    # Leitura das planilhas
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

      # Criar uma nova planilha para o resultado da comparação
    result_df = pd.DataFrame(columns=df1.columns)

    # Comparar os valores da coluna A e gravar o resultado na nova planilha
    for index, row1 in df1.iterrows():
        for index2, row2 in df2.iterrows():
            if row1['Coluna A'] == row2['Coluna A'] and (row1['Coluna B'] == row2['Coluna C'] or row1['Coluna B'] == row2['Coluna D']  ) :
                result_df = result_df.append(row1)

    # Gravar o resultado da comparação na nova planilha
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Resultado da Comparação"

    for row in dataframe_to_rows(result_df, index=False, header=True):
        sheet.append(row)

    # Salvar o arquivo Excel
    workbook.save(output_file)
    print(f"O resultado foi gravado na aba 'Resultado da Comparação' do arquivo '{output_file}'.")


if __name__ == "__main__":
    # Substitua "caminho/para/planilha1.xlsx", "caminho/para/planilha2.xlsx" e "caminho/para/saida.xlsx" pelos caminhos corretos das planilhas que deseja comparar e do arquivo de saída.
    file1_path = "C:\\Users\\walter.oliveira\\Documents\\Manuais_Bichara\\RD_Sisjuri_Previa.xlsx"
    file2_path = "C:\\Users\\walter.oliveira\\Documents\\Manuais_Bichara\\SisJuri.xlsx"
    output_file_path = "C:\\Users\\walter.oliveira\\Documents\\Manuais_Bichara\\saida.xlsx"

    compare_excel_sheets(file1_path, file2_path, output_file_path)
