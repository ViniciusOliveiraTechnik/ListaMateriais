# Libs de manipulação para o Excel
import openpyxl as op
from openpyxl.styles import Alignment, Font
from openpyxl.styles.borders import Border, Side

# Libs variadas
from math import ceil
import pandas as pd

# Classe de estilos para o Excel
class Styles:
    def __init__(self):
        
        self.titleFont = Font(name='Arial', bold=True, size=8) # Font para titúlos
        self.centerAlignment = Alignment(vertical='center', horizontal='center') # Alinhamento no meio
        self.standardBorder = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            bottom=Side(style='thin'),
            top=Side(style='thin')
        ) # Borda padrão
        self.standardFont = Font(name='Arial', size=8) # Fonte padrão
        self.centerLeftFont = Alignment(vertical='center', horizontal='left', wrap_text=True) # Fonte centralizada ao meio e esquerda

    # Método que aplica borda
    def apply_border(self, start, end, work_sheet):

        # Aplica um intervalo entre células
        for row in work_sheet[start:end]:

            # Para cada uma dentro desse intervalo aplique a borda
            for cell in row:
                
                # Tenta aplicar a borda para a célula
                try:
                    cell.border = self.standardBorder
                
                # Retorna erro inesperado
                except Exception as err:
                    return {'status': False, 'error': f'Erro ao aplicar estilo de borda: "{str(err)}"'}
                
# Classe para o tratamento dos parafusos
class Screws:
    def __init__(self):

        # Dataframe utilizado para manipulação dos parafusos
        self.baseDF= pd.DataFrame({
        'classe': [150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150, 150],
        'diametro_nominal': [
            '1/2"', '3/4"', '1"', '1 1/4"', '1 1/2"', '2"', '2 1/2"', '3"', '4"', '6"', '8"', '10"', '12"', '14"', '16"', '18"', '20"', '24"'
        ],
        'parafusos_quantidade': [4, 4, 4, 4, 4, 4, 4, 4, 8, 8, 8, 12, 12, 12, 16, 16, 20, 20],
        'parafusos_diametro': [
            '1/2"', '1/2"', '1/2"', '1/2"', '1/2"', '5/8"', '5/8"', '5/8"', '5/8"', '3/4"', '3/4"', '7/8"', '7/8"', '1"', '1"', '1 1/8"', '1 1/8"', '1 1/4"'
        ],
        'comprimento': [
            '2"', '2 1/4"', '2 1/4"', '2 1/2"', '2 1/2"', '2 3/4"', '3"', '3 1/4"', '3 1/4"', '3 1/2"', '3 3/4"', '4"', '4 1/4"', '4 1/2"', '4 3/4"', '5"', '5 1/2"', '6"'
        ]
        })

    # Método para obter todos os parafusos do dataframe
    def get_screws(self, DataFrame):

        # Tentativa de pegar os parafusos
        try:

            # Filtra dataframe para obter apenas a descrição que seja igual a Flange
            screwsDF = DataFrame[DataFrame['Long Description (Size)'].str.startswith('FLANGE')][['Long Description (Size)', 'Size', 'Spec', 'Fixed Length']]

            # Filtra através do banco de dados

            for column in self.baseDF.columns:

                # Cria um lookup em todas as colunas dos dataframe
                screwsDF[column] = screwsDF['Size'].map(    
                    lambda x: self.baseDF[self.baseDF['diametro_nominal'] == str(x)][column].iloc[0]
                )

            # Atualizar descrições e tamanhos
            screwsDF['Long Description (Size)'] = screwsDF.apply(
                lambda row: f"PARAFUSO {row['parafusos_diametro']} X {row['comprimento']}", axis=1
            )

            # Multiplica a quantidade de flanges pela quantidade de parafusos
            screwsDF['Fixed Length'] *= screwsDF['parafusos_quantidade']

            # Cria um parafuso
            screwsDF['Size'] = screwsDF.apply(
                lambda row: f"{row['parafusos_diametro']}x{row['comprimento']}", axis=1
            )

            # Retorna todos os parafusos
            return {'status': True, 'data': screwsDF} 

        except KeyError as err:
            return {'status': False, 'error': f'KeyError relacinado à "{str(err)}" '}
        
        except Exception as err:
            return {'status': False, 'error': f'Erro Inesperado: {str(err)}'}

# Classe para extração de dados do excel
class ExcelExtract:
    def __init__(self):
        pass
    
    # Método para concatenar diversos DataFrames
    def concat(self, *args):
        # Verifica se os argumentos foram passados
        if not args:
            return {'status': False, 'errror': f'Nenhum dado foi passado para concatenar: {args}'} # Retorna erro
        
        # For para cada argumento validando se ele é do tipo DataFrame
        for arg in args:
            if not isinstance(arg, pd.DataFrame):
                return {'status': False, 'error': f'O argumento "{arg}" não é do tipo DataFrame'} # Retorna erro
        
        # Define o DataFrame geral como a concatenação dos passados
        try:
            mainDF = pd.concat(args, ignore_index=True)
            return {'status': True, 'data': mainDF} # Retorna dados corretamente
        
        # Tratamento de erros
        except Exception as err:
            return {'status': False, 'error': f'Erro Inesperado: {str(err)}'}
            
    # Método para somar valores únicos de cada dataframe
    def sum_unique(self, dataframe):

        # Tentativa de executar o programa
        try:

            # Agrupa os dados do dataframe que possuem dados semelhantes
            mainDF = dataframe.groupby(['Long Description (Size)', 'Size', 'Spec'], as_index=False)['Fixed Length'].sum()
            return {'status': True, 'data': mainDF} # Retorna status positivo e dados

        # Tratamento de erros
        except ValueError as err:
            return {'status': False, 'error': f'Erro entre os tipos de valores: {str(err)}'}
        
        except KeyError as err:
            return {'status': False, 'error': f'Erro de chaves da tabela: {str(err)}'}

        except Exception as err:
            return {'status': False, 'error': f'Erro Inesperado: {str(err)}'}
        
    # Método para somar valores únicos de cada dataframe
    def count_unique(self, dataframe):

        # Tentativa de executar o programa
        try:

            # Agrupa os dados do dataframe que possuem dados semelhantes
            mainDF = dataframe.groupby(['Long Description (Size)', 'Size', 'Spec']).size().reset_index(name='Fixed Length')

            # Cria objeto da classe de parafusos
            screws = Screws()
            response = screws.get_screws(mainDF)

            # Verifica se a response é verdadeira
            if response['status']:
                screwsDF = response['data'] # Obtem os dados do dicionário

                response = self.sum_unique(screwsDF) # Segunda response de método, para formatar os parafusos

                # Verifica o status de execução é positivo
                if response['status']:  
                    screwsDF = response['data'] # Obtem todos os dados do dicionario
                    mainDF = pd.concat([mainDF, screwsDF], ignore_index=True) # Concatena os dados do mainDF e do screwDF

                    return {'status': True, 'data': mainDF} # Retorna status positivo e dados
                
                return {'status': False, 'error': response['error']} # Retorna erro caso algum ocorra

            return {'status': False, 'error': response['error']} # Retorna erro caso algum ocorra
        
        # Tratamento de erros
        except ValueError as err:
            return {'status': False, 'error': f'Erro entre os tipos de valores: {str(err)}'}
        
        except KeyError as err:
            return {'status': False, 'error': f'Erro de chaves da tabela: {str(err)}'}

        except Exception as err:
            return {'status': False, 'error': f'Erro Inesperado: {str(err)}'}

    # Método para ler todos os arquivos
    def read_all_files(self, files, sheet_name):

        # Erro caso os parâmetros não sejam satisfeitos
        if not sheet_name or not files:
            return {'status': False, 'error': 'Os parâmetros não foram fornecidos'}

        mainDF = pd.DataFrame() # Cria dataframe vazio

        # Aplica um for para cada arquivo selecionado
        for file in files:
            
            # Tentativa de ler o arquivo com pandas
            try:
                tempDF = pd.read_excel(file, sheet_name=sheet_name)
                mainDF = pd.concat([mainDF, tempDF], ignore_index=True) # Concatena o dataframe temporário com o dataframe principal
            
            # Tratamento de erros
            except KeyError as err:
                return {'status': False, 'error': f'Erro de Chaves: {str(err)}'}
            
            except FileNotFoundError as err:
                return {'status': False, 'error': f'Não foi possível encontrar o arquivo: {str(err)}'}
            
            except Exception as err:
                return {'status': False, 'error': f'Erro Inesperado, solicite o suporte da equipe de TI: {str(err)}'}

        # Tentativa de filtragem do dataframe com as colunas especificadas
        try:
            mainDF = mainDF[['Long Description (Size)', 'Spec', 'Size', 'Fixed Length']]
        
        # Caso Fixed Length não exista dentro desse DataFrame tenta-se outra abordagem
        except KeyError as err:
            try: # Tentativa de criar DataFrame com as colunas especificadas
                mainDF = mainDF[['Long Description (Size)', 'Spec', 'Size']]

            # Tratamento de erro para todos os casos
            except KeyError as err:
                return {'status': False, 'error': f'Erro de Chaves: {str(err)}'}

            except Exception as err:
                return {'status': False, 'error': f'Erro Inesperado, solicite o suporte da equipe de TI: {str(err)}'}

        # Deleta todas as linhas que possuirem valores do tipo NaN
        mainDF = mainDF.dropna().reset_index(drop=True)

        # Aplica a cada SPEC a verificação se ela possui 0 antes de algum valor numérico
        mainDF['Spec'] = mainDF['Spec'].map(lambda x: str(x).replace('0', '') if str(x).startswith('0') and len(str(x)) > 1 
        else str(x).upper())

        # Verifica se o nome da Sheet é Pipe (tubulação)
        if 'pipe' == str(sheet_name).lower():
            
            # Tentativa de agrupar valores únicos
            try:
                response = self.sum_unique(mainDF) # Resposta de execução do método

                # Verifica se o status é verdadeiro
                if response['status']:
                    mainDF = response['data'] # Obtem os dados 
                    mainDF['Categorie'] = 'm' # Define a categoria como metro

                    # Retorna status verdadeiro e dados
                    return {'status': True, 'data': mainDF}
                
                # Retorna erro caso algum ocorra
                return {'status': False, 'error': response['error']}
            
            # Tratamento de erros
            except KeyError as err:
                return {'status': False, 'error': f'Erro de Chaves: {str(err)}'}
            
            except ValueError as err:
                return {'status': False, 'error': f'Erro de Valores: {str(err)}'}
            
            except Exception as err:
                return {'status': False, 'error': f'Erro Inesperado, solicite o suporte da equipe de TI: {str(err)}'}

        # Tentativa de filtrar equipamentos
        try:
            mainDF = mainDF[~mainDF['Long Description (Size)'].str.startswith('TUBO')] # Filtra todos os valores exceto tubulações
        
            response = self.count_unique(mainDF) # Resposta do método de agrupamento por contagem única

            # Verifica se o status é positivo
            if response['status']:
                mainDF = response['data'] # Obtem dados
                mainDF['Categorie'] = 'pç' # Define a categoria como peça

                return {'status': True, 'data': mainDF} # Retorna status positivo e dados
            
            # Retorna erro caso algum ocorra
            return {'status': False, 'error': response['error']}

        # Tratamento de erros
        except KeyError as err: 
            return {'status': False, 'error': f'Erro de Chaves: {str(err)}'}
            
        except ValueError as err:
            return {'status': False, 'error': f'Erro de Valores: {str(err)}'}
            
        except Exception as err:
            return {'status': False, 'error': f'Erro Inesperado, solicite o suporte da equipe de TI: {str(err)}'}
