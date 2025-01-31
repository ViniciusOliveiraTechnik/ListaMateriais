# Libs de manipulação para o Excel
import openpyxl as op
from openpyxl.styles import Alignment, Font
from openpyxl.styles.borders import Border, Side

# Libs variadas
import pandas as pd
from math import ceil

# Classe de estilos para o Excel
class Styles:
    def __init__(self):
        
        self.title_font = Font(name='Arial', bold=True, size=8) # Font para titúlos
        self.title_font_header = Font(name='Arial', bold=True, size=10) # Font para titúlos
        self.center_align = Alignment(vertical='center', horizontal='center') # Alinhamento no meio
        self.center_wrap_align = Alignment(vertical='center', horizontal='center', wrap_text=True)
        self.standard_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            bottom=Side(style='thin'),
            top=Side(style='thin')
        ) # Borda padrão
        self.standard_font = Font(name='Arial', size=8) # Fonte padrão
        self.standard_font_header = Font(name='Arial', size=10) # Fonte padrão
        self.center_left_align = Alignment(vertical='center', horizontal='left', wrap_text=True) # Fonte centralizada ao meio e esquerda
        self.top_rigth_align = Alignment(vertical='top', horizontal='right', wrap_text=True) # Alinhamento topo e direita

    # Método que aplica borda
    def apply_border(self, start, end, work_sheet):

        # Aplica um intervalo entre células
        for row in work_sheet[start:end]:
            # Para cada uma dentro desse intervalo aplique a borda
            for cell in row:
                # Tenta aplicar a borda para a célula
                try:
                    cell.border = self.standard_border
                
                # Retorna erro inesperado
                except Exception as err:
                    return {'status': False, 'error': f'Erro ao aplicar estilo de borda: "{str(err)}"'}        
    
    # Método para criar o cabeçalho do arquivo
    def create_header(self, worksheet):

        # Bloco superior esquerdo
        worksheet.merge_cells('A1:C4')
        
        # Bloco Technik
        worksheet.merge_cells('A5:C8')

        worksheet['A5'] = '=IMAGEM("https://technikgrupo.sharepoint.com/sites/transfer/Logos/Technik.png")' # Adicionando conteúdo de Imagem
        worksheet['A5'].alignment = self.center_align  # Alinhamento ao centro

        # Legenda Technik
        worksheet.merge_cells('A9:C10')

        worksheet['A9'] = 'CENTRO DE ENGENHARIA TECHNIK'
        worksheet['A9'].alignment = self.center_wrap_align # Alinhamento ao centro com quebra de linha
        worksheet['A9'].font = self.standard_font # Fonte padrão de texto

        # Título superior     
        worksheet.merge_cells('D1:G2')

        worksheet['D1'] = 'LISTA DE MATERIAIS'
        worksheet['D1'].alignment = self.center_align
        worksheet['D1'].font = self.title_font_header

        # Informações do cliente
        
        # Cliente
        worksheet['D3'] = 'CLIENTE'
        worksheet['D3'].alignment =  self.center_left_align
        worksheet['D3'].font = self.standard_font

        worksheet['D4'] = 'Client'
        worksheet['D4'].alignment =  self.center_left_align
        worksheet['D4'].font = self.standard_font

        worksheet.merge_cells('E3:L4')

        # Unidade
        worksheet['D5'] = 'UNIDADE'
        worksheet['D5'].alignment =  self.center_left_align
        worksheet['D5'].font = self.standard_font

        worksheet['D6'] = 'Plant'
        worksheet['D6'].alignment =  self.center_left_align
        worksheet['D6'].font = self.standard_font

        worksheet.merge_cells('E5:L6')

        # Área
        worksheet['D7'] = 'ÁREA'
        worksheet['D7'].alignment =  self.center_left_align
        worksheet['D7'].font = self.standard_font

        worksheet['D8'] = 'Area'
        worksheet['D8'].alignment =  self.center_left_align
        worksheet['D8'].font = self.standard_font

        worksheet.merge_cells('E7:L8')

        # Título
        worksheet['D9'] =  'Título'
        worksheet['D9'].alignment =  self.center_left_align
        worksheet['D9'].font = self.standard_font

        worksheet['D10'] =  'Title'
        worksheet['D10'].alignment =  self.center_left_align
        worksheet['D10'].font = self.standard_font

        worksheet.merge_cells('E9:l10')

        worksheet['E9'] = 'LISTA DE MATERIAIS DE TUBULAÇÃO'
        worksheet['E9'].alignment =  self.center_align
        worksheet['E9'].font = self.standard_font_header

        # Número do fluxograma
        worksheet.merge_cells('H1:H2')

        worksheet['H1'] =  'Nº'  # Label descritiva
        worksheet['H1'].alignment =  self.top_rigth_align
        worksheet['H1'].font = self.standard_font

        worksheet.merge_cells('I1:L2') # Campo para inserir valores
                
        row_index = 10

        return row_index
    
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
    # def get_screws(self, DataFrame):

        #

        # # Tentativa de pegar os parafusos
        # try:
        #     # Filtra dataframe para obter apenas a descrição que seja igual a Flange
        #     screwsDF = DataFrame[DataFrame['Long Description (Family)'].str.startswith('FLANGE')][['Long Description (Family)', 'Size', 'Spec', 'Fixed Length']]

        #     # Filtra através do banco de dados
        #     for column in self.baseDF.columns:
        #         # Cria um lookup em todas as colunas dos dataframe
        #         screwsDF[column] = screwsDF['Size'].map(    
        #             lambda x: self.baseDF[self.baseDF['diametro_nominal'] == str(x)][column].iloc[0])

        #     # Atualizar descrições e tamanhos
        #     screwsDF['Long Description (Family)'] = screwsDF.apply(
        #         lambda row: f"PARAFUSO {row['parafusos_diametro']} X {row['comprimento']}", axis=1)
            
        #     # Multiplica a quantidade de flanges pela quantidade de parafusos
        #     screwsDF['Fixed Length'] *= screwsDF['parafusos_quantidade']

        #     # Cria um parafuso
        #     screwsDF['Size'] = screwsDF.apply(
        #         lambda row: f"{row['parafusos_diametro']}x{row['comprimento']}", axis=1)

        #     # Retorna todos os parafusos
        #     return {'status': True, 'data': screwsDF} 

        # except KeyError as err:
        #     return {'status': False, 'error': f'KeyError relacinado à "{str(err)}" '}
        # except Exception as err:
        #     return {'status': False, 'error': f'Erro Inesperado: {str(err)}'}

# Classe para extração de dados do excel
class ExcelExtract:
    def __init__(self):
        pass
    
    # Método para formatar o valor do Fixed Length
    def ceil_format(self, value: float, categorie: str = 'm', extra_percent: float = 0):
        # Define value como new_value
        new_value = float(value)

        # Verifica se a categioria é metro
        if categorie == 'm':
            new_value /= 1000 # Altera a escala de centímetro para metro
        
        # Adiciona o valor extra de percentual
        new_value = ceil(new_value + (new_value * extra_percent))

        return float(new_value) # Retorna o valor em formato de float
    
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
            return {'status': True, 'data': mainDF.sort_values(by=['Long Description (Family)', 'Size', 'Spec'])} # Retorna dados corretamente
        
        # Tratamento de erros
        except Exception as err:
            return {'status': False, 'error': f'Erro Inesperado: {str(err)}'}
            
    # Método para somar valores únicos de cada dataframe
    def sum_unique(self, dataframe):
        # Tentativa de executar o programa
        try:
            # Agrupa os dados do dataframe que possuem dados semelhantes
            mainDF = dataframe.groupby(['Long Description (Family)', 'Size', 'Spec'], as_index=False)['Fixed Length'].sum()
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
            mainDF = dataframe.groupby(['Long Description (Family)', 'Size', 'Spec']).size().reset_index(name='Fixed Length')
            return {'status': True, 'data': mainDF} # Retorna status positivo e dados
        
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
            return {'status': False, 'error': 'Os parâmetros não foram fornecidos para a função "read_all_files"'}

        mainDF = pd.DataFrame() # Cria dataframe vazio

        # Aplica um for para cada arquivo selecionado
        for file in files:
            try: 
                tempDF = pd.read_excel(file, sheet_name=sheet_name) # Lê arquivo excel na folha selecionada
                mainDF = pd.concat([mainDF, tempDF], ignore_index=True) # Concatena o dataframe temporário com o dataframe principal
            
            # Tratamento de erros
            except KeyError as err:
                return {'status': False, 'error': f'Não foi possível encontrar a planilha desejada: "{str(err)}"'}
            except FileNotFoundError as err:
                return {'status': False, 'error': f'Não foi possível encontrar o arquivo: "{str(err)}"'}
            except Exception as err:
                return {'status': False, 'error': f'Erro Inesperado, solicite o suporte da equipe de TI: "{str(err)}"'}

        try:
            # Filtra o DataFrame para calcular as tubulações
            mainDF = mainDF[['Long Description (Family)', 'Status', 'Spec', 'Size', 'Fixed Length']]
        
        # Caso Fixed Length não exista dentro desse DataFrame tenta-se outra abordagem
        except KeyError as err:
            try: # Tentativa de criar DataFrame com as colunas especificadas
                mainDF = mainDF[['Long Description (Family)', 'Status', 'Spec', 'Size']]

            # Tratamento de erro para todos os casos
            except KeyError as err:
                return {'status': False, 'error': f'Erro ao encontrar as colunas desejadas: "{str(err)}"'}
            except Exception as err:
                return {'status': False, 'error': f'Erro Inesperado, solicite o suporte da equipe de TI: "{str(err)}"'}

        # Filtrando o MainDF
        try:
            mainDF = mainDF.dropna().reset_index(drop=True) # Deleta todas as linhas que possuirem valores do tipo NaN
            mainDF = mainDF[mainDF['Status'] == 'New'] # Filtra apenas o status new

            mainDF['Spec'] = (
                mainDF['Spec']
                .astype(str)
                .str.upper()
                .str.replace(r'(^\bSPEC\b|\b0\d+)', '', regex=True)
                .str.strip()
                )


        except KeyError as err:
            return {'status': False, 'error': f'Erro ao encontrar colunas: {str(err)}'}
        except Exception as err:
            return {'status': False, 'error': f'Erro Inesperado, solicite o suporte da equipe de TI: "{str(err)}"'}

        # Verifica se o nome da Sheet é Pipe (tubulação)
        if 'pipe' == str(sheet_name).lower():
            try:
                # Somando valores únicos
                response = self.sum_unique(mainDF)

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

        # Filtra apenas os equipamentos
        try:
            mainDF = mainDF[~mainDF['Long Description (Family)'].str.upper().str.contains(r'^(TUBO|PIPE|PARAFUSO|FLANGE)', regex=True)] # Filtra todos os valores exceto tubulações e flanges
        
            # Conta e define valores únicos
            response = self.count_unique(mainDF)

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
