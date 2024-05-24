from flask import Flask, request, jsonify
import requests
from bs4 import BeautifulSoup
import openpyxl

app = Flask(__name__)

def consultar_debito(cda):
    url = f'https://www2.fazenda.mg.gov.br/sol/ctrl/SOL/DIVATIV/SERVICO_001?ACAO=VISUALIZAR={cda}'
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Aqui você pode implementar a lógica para extrair as informações desejadas do HTML
    # e retornar essas informações como um dicionário.

    # Exemplo hipotético:
    informacoes = {
        'nome_devedor': soup.find('span', {'id': 'nomeDevedor'}).text,
        'valor_debito': soup.find('span', {'id': 'valorDebito'}).text,
        # ... outras informações ...
    }

    return informacoes

@app.route('/consulta', methods=['GET'])
def api_consulta():
    cda = request.args.get('cda')
    if not cda:
        return jsonify({'error': 'Parâmetro CDA ausente.'}), 400

    try:
        informacoes = consultar_debito(cda)

        # Gravar as informações em um arquivo Excel
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        for key, value in informacoes.items():
            sheet.append([key, value])

        filename = f'informacoes_{cda}.xlsx'
        workbook.save(filename)

        return jsonify(informacoes)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)








