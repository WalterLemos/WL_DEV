from typing import Dict

def main():
    import json
    import sys

    jobs = json.load(open("config_export.json"))
    files = []
    if len(sys.argv) > 1:
        files = f'{sys.argv[1]}'.split(',')

    if files:
        for file in files:
            for job in jobs:
                if file in job.keys():
                    exportar(file, job[file])
    else:
        for job in jobs:
            for file in job.keys():
                exportar(file, job[file])

    return

def get_cursor():
    import cx_Oracle
    import os

    ORACM_USER = os.getenv('ORACM_USER') 
    ORACM_PWD = os.getenv('ORACM_PWD')
    ORACM_LIB = os.getenv('ORACLE_HOME')
    
    dsn = cx_Oracle.makedsn('192.168.0.251', 1521, 'mega')
    try:
        cx_Oracle.init_oracle_client(lib_dir=ORACM_LIB)
    except:
        pass
    
    return cx_Oracle.connect(ORACM_USER, ORACM_PWD, dsn).cursor()

def exportar(nome_arquivo: str, abas: Dict):
    from datetime import datetime
    import xlsxwriter

    cursor = get_cursor()
    with xlsxwriter.Workbook(nome_arquivo) as wb:
        for aba in abas:
            planilha = wb.add_worksheet(aba['sheet_name'])
            sql = open(aba['sql'], 'r')
            
            # print(''.join(sql.readlines()))
            cursor.execute(''.join(sql.readlines()))
            registros = cursor.fetchall()

            # Cabecalho
            for coluna, campo in enumerate(cursor.description):
                planilha.write(0, coluna, campo[0])
                planilha.set_column(coluna, coluna, width=len(campo)*2)

            # Dados
            for linha, registro in enumerate(registros):
                for coluna, campo in enumerate(registro):
                    fmt = None
                    if type(campo) == datetime:
                        fmt = wb.add_format({'num_format': 'dd/mm/yyyy'})
                    if type(campo) == float:
                        fmt = wb.add_format({'num_format': '#,##0.00;-#,##0.00'})
                    planilha.write(linha+1, coluna, campo, fmt)


if __name__ == '__main__':
    main()