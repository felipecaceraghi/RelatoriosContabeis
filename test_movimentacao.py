import os
import sys
sys.path.append(r'C:\Users\estatistica007\Documents\nextProjects\RelatoriosContabeis\scripts')
from file_renamer import mover_arquivo_para_destino

# Criar um arquivo de teste
with open('teste.txt', 'w') as f:
    f.write('Arquivo de teste')

# Testar movimentação
resultado = mover_arquivo_para_destino('teste.txt', '124', '2025', '08')
print(f'Resultado da movimentação: {resultado}')

# Verificar se o arquivo foi movido
if os.path.exists('teste.txt'):
    print('Arquivo ainda existe no local original')
else:
    print('Arquivo foi movido com sucesso')
