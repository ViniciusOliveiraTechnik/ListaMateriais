import re

string = """
SPEC 02
SPEC 16
CS150
ALLTECH 07
"""

pattern = r'(SPEC|0)'

string = re.sub(string=string, pattern=pattern, repl='').strip()

print(string)

# import pandas as pd

# path = r'c:\Users\Vinicius\Downloads\TUBULAÇÃO_OP02-Piping and Equipment.xlsx'

# pipe = pd.read_excel(path, 'Pipe')

# pipe = pipe[df['Satatus']]