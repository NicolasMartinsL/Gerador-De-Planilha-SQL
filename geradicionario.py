#
#Codigo Desenvolvido por Nicolas Martins Lorena
#
from openpyxl import load_workbook
import pandas as pd
import json

class ConverterPlanilhas:

    def __init__(self):
        # Carregar a planilha existente
        self.workbook = load_workbook("doismilturismo.xlsx")
        self.sheet = self.workbook.active  # Selecionar a primeira planilha

        # Renomear colunas
        self.sheet['A1'] = 'id'
        self.sheet['B1'] = 'Identificação'
        self.sheet['C1'] = 'Exame/Serviço'
        self.sheet['D1'] = 'Valor interno'

    def gera_dicionario(self):
        # Nome do arquivo Excel
        arquivo_excel = 'tabela.xlsx'

        # Lê a planilha do Excel
        df = pd.read_excel(arquivo_excel, sheet_name=0)
        print("DataFrame carregado:")
        print(df)

        # Converte os dados para uma lista de dicionários
        dados = df.to_dict(orient='records')

        # Filtra os dados
        dados_filtrados = [
            {
                'ID': item['ID'],  # Certifique-se de que 'ID' existe no arquivo Excel
                'Identificação': item['Identificacao'],
                'Exame/Serviço': item['Exame/Serviço']
            }
            for item in dados
        ]

        # Salva os dados filtrados em um arquivo JSON
        arquivo_json = 'dicionarioFinal.json'
        with open(arquivo_json, 'w', encoding='utf-8') as json_file:
            json.dump(dados_filtrados, json_file, ensure_ascii=False, indent=4)

        print(f'Dados salvos em {arquivo_json}')

    def obter_informacoes(self, identificacao_procurado, arquivo):
        with open(arquivo, 'r', encoding='utf-8') as f:  # Usando utf-8 para evitar erros de caracteres
            dados = json.load(f)

            # Iterar pela lista para encontrar o item correspondente
            for item in dados:
                if item.get('Identificação') == identificacao_procurado:
                    return item
            return None

# Exemplo de uso:
arquivo = 'dicionarioFinal.json'
identificacao_procurado = "CLORETO"
teste = ConverterPlanilhas()
teste.gera_dicionario()
resultado = teste.obter_informacoes(identificacao_procurado, arquivo)

if resultado:
    id = resultado.get('ID')
    classificacao = resultado.get('Exame/Serviço')
    print(f"O ID de {identificacao_procurado} é {id} e é um {classificacao}.")
else:
    print(f"O identificador {identificacao_procurado} não foi encontrado.")
