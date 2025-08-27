#!/usr/bin/env python3
"""
Script de teste para verificar detecção de tipos de arquivo
"""
import os
import sys

# Adicionar o diretório atual ao path para importar file_renamer
sys.path.append(os.path.dirname(__file__))

def test_file_type_detection():
    print('Testando detecção de tipos de arquivo:')

    test_files = [
        'razao_emp_124_2025-07-01_a_2025-07-31_20250827_114100.xlsx',
        'DRE_EMP124_20250101_a_20250731_EN_20250827_114100.pdf',
        'Balancete_124_2025-07-01_a_2025-07-31_20250827_114100.xlsx',
        'Comparativo_Movimento_124_20250701_20250731_20250827_114100.xlsx'
    ]

    for filename in test_files:
        tipo = None
        if 'Balancete' in filename or 'balancete' in filename:
            tipo = 'balancete'
        elif 'Comparativo' in filename or 'comparativo' in filename:
            tipo = 'comparativo'
        elif 'DRE' in filename or 'dre' in filename:
            tipo = 'dre'
        elif 'razao' in filename or 'Razao' in filename:
            tipo = 'razao'
        else:
            tipo = None

        print(f'{filename} -> {tipo}')

if __name__ == '__main__':
    test_file_type_detection()
