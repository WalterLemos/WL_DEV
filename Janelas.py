from bs4 import BeautifulSoup

# Conteúdo HTML
html_content = """
<td class="rich-mpnl-body" valign="top"><br><br><label>Situação Protesto:</label> Enviado para cartório de protesto<br><label>Comarca:</label> SAO PAULO<br><label>N° Protocolo:</label> 369<br><label>Data de abertura:</label> 17/06/2024<br><br><br><table class="rich-table" id="consultaDebitoForm:tabelaDadosCartorio" style="margin-top: 10px; width: 780px !important; margin-bottom: 10px; margin-left: auto; margin-right: auto;" border="0" cellpadding="0" cellspacing="0"><colgroup span="7"></colgroup><thead class="rich-table-thead"><tr class="rich-table-header"><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1284" colspan="7" style="text-align: left;">Dados do cartório de protesto</th></tr><tr class="rich-table-header-continue"><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1286" style="width: 160px">Nome</th><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1288" style="width: 120px">Endereço</th><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1290" style="width: 120px">Bairro</th><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1292" style="width: 120px">Localidade</th><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1294" style="width: 60px">CEP</th><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1296" style="width: 100px">E-mail</th><th class="rich-table-headercell" id="consultaDebitoForm:tabelaDadosCartorio:j_id1298" style="width: 80px">Telefone</th></tr></thead><tbody id="consultaDebitoForm:tabelaDadosCartorio:tb"><tr class="rich-table-row rich-table-firstrow"><td class="rich-table-cell" id="consultaDebitoForm:tabelaDadosCartorio:0:j_id1300" style="text-align: center;">5.o TABELIONATO DE PROTESTO DE LETRAS E TITULOS</td><td class="rich-table-cell" id="consultaDebitoForm:tabelaDadosCartorio:0:j_id1302" style="text-align: center;">RUA DA GLORIA, 162</td><td class="rich-table-cell" id="consultaDebitoForm:tabelaDadosCartorio:0:j_id1304" style="text-align: center;">LIBERDADE</td><td class="rich-table-cell" id="consultaDebitoForm:tabelaDadosCartorio:0:j_id1306" style="text-align: center;">SAO PAULO</td><td class="rich-table-cell" id="consultaDebitoForm:tabelaDadosCartorio:0:j_id1308" style="text-align: center;">01510000</td><td class="rich-table-cell" id="consultaDebitoForm:tabelaDadosCartorio:0:j_id1310" style="text-align: center;"></td><td class="rich-table-cell" id="consultaDebitoForm:tabelaDadosCartorio:0:j_id1312" style="text-align: center;">1132423143</td></tr></tbody></table></td>
"""

# Analise o conteúdo HTML
soup = BeautifulSoup(html_content, 'html.parser')

# Encontre todos os valores de texto
text_values = [value.text.strip() for value in soup.find_all('td', class_='rich-table-cell')]

# Imprima os valores de texto
for value in text_values:
    print(value)

    
