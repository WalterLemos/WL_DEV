import pandas as pd
import pyodbc

def consulta_access_e_grava_excel():
    # Configurações da conexão com o Access
    access_driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    access_file = 'C:\\Users\\walter.oliveira\\Documents\\Database1.accdb'  # Substitua pelo caminho do seu arquivo Access
    conn_str = f'DRIVER={access_driver};DBQ={access_file};'

    # Consulta SQL que deseja executar
    consulta_sql = '''
        SELECT DISTINCT SGC.[Cód# Grupo de Empresa], SGC.[Descrição do Grupo de Empresa] AS EMPRESA 
        FROM SISJURI_GRUPO_CLIENTES AS SGC, [rd-bichara-advogados] AS RD
        WHERE    
        SGC.[Descrição do Grupo de Empresa] = RD.[Empresa] OR 
        SGC.[Razão Social] =  RD.[Empresa];
    '''
# Conectar ao banco de dados Access
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    except Exception as e:
        print(f'Erro ao conectar ao banco de dados Access: {e}')
        return

    # Executar a consulta e obter os resultados
    try:
        cursor.execute(consulta_sql)
        rows = cursor.fetchall()
        col_names = [column[0] for column in cursor.description]
        df_resultado = pd.DataFrame(rows, columns=col_names)
    except Exception as e:
        print(f'Erro ao executar a consulta: {e}')
        conn.close()
        return

    # Gravar os resultados no Excel
    try:
        nome_arquivo_excel = 'resultado_consulta.xlsx'  # Substitua pelo nome do arquivo Excel que deseja criar
        df_resultado.to_excel(nome_arquivo_excel, index=False)
        print(f'Resultado da consulta gravado com sucesso em {nome_arquivo_excel}')
    except Exception as e:
        print(f'Erro ao gravar os resultados no Excel: {e}')
    finally:
        conn.close()

if __name__ == "__main__":
    consulta_access_e_grava_excel()
