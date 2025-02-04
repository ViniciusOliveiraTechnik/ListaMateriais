required_columns = ['Long Description (Family)', 'Spec', 'Size', 'Fixed Length']

index = required_columns.index('Fixed Length')
required_columns.pop(index)
print(required_columns)