import os
import sys
sys.path.append(r'C:\Users\estatistica007\Documents\nextProjects\RelatoriosContabeis\scripts')
from file_renamer import encontrar_pasta_cliente

# Testar com alguns códigos de empresa
testes = ['124', '2018']
for codi_emp in testes:
    pasta = encontrar_pasta_cliente(codi_emp)
    print(f'Empresa {codi_emp}: {pasta}')
