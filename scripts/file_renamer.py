import pyodbc
# Função utilitária para buscar o nome da empresa no banco
def buscar_nome_empresa(codi_emp: str, conn_str: str) -> str:
    """
    Busca o nome (apel_emp) da empresa no banco de dados pelo código.
    conn_str: string de conexão ODBC obrigatória.
    """
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute("select g.apel_emp from bethadba.geempre g where g.codi_emp = ?", codi_emp)
            row = cursor.fetchone()
            if row:
                nome = row[0].strip()
                # Remove todas as ocorrências do código da empresa em qualquer parte do nome
                nome = nome.replace(str(codi_emp), '').strip()
                return nome
    except Exception as e:
        print(f"Erro ao buscar nome da empresa: {e}")
    return str(codi_emp)
import shutil

# Caminho base da rede
REDE_BASE = r"R:\\Acesso Digital"

def encontrar_pasta_cliente(codi_emp: str, base_path: str = REDE_BASE) -> str:
    """
    Busca a pasta do cliente pelo código dentro do diretório base.
    Retorna o caminho completo da pasta do cliente ou None se não encontrar.
    """
    try:
        for nome in os.listdir(base_path):
            if nome.startswith(str(codi_emp) + ' ' ) or nome.startswith(str(codi_emp) + '-') or nome.startswith(str(codi_emp) + '_') or nome.startswith(str(codi_emp)):
                if nome.split()[0] == str(codi_emp):
                    return os.path.join(base_path, nome)
    except Exception as e:
        print(f"Erro ao buscar pasta do cliente: {e}")
    return None

def montar_caminho_destino(pasta_cliente: str, ano: str, mes: str) -> str:
    """
    Monta o caminho final para os relatórios contábeis do cliente.
    """
    caminho = os.path.join(
        pasta_cliente,
        "02 - Contábil",
        ano,
        "01 - Fechamento Contábil",
        mes,
        "03 - Relatórios contábeis",
        "GERADO PELO ROBO"
    )
    return caminho

def mover_arquivo_para_destino(arquivo_path: str, codi_emp: str, ano: str, mes: str):
    """
    Move o arquivo para o destino correto na rede, criando as pastas se necessário.
    """
    pasta_cliente = encontrar_pasta_cliente(codi_emp)
    if not pasta_cliente:
        print(f"Pasta do cliente {codi_emp} não encontrada em {REDE_BASE}")
        return False
    destino = montar_caminho_destino(pasta_cliente, ano, mes)
    try:
        os.makedirs(destino, exist_ok=True)
        nome_arquivo = os.path.basename(arquivo_path)
        destino_final = os.path.join(destino, nome_arquivo)
        shutil.move(arquivo_path, destino_final)
        print(f"Arquivo movido para: {destino_final}")
        return True
    except Exception as e:
        print(f"Erro ao mover arquivo para destino: {e}")
        return False
#!/usr/bin/env python3
"""
Módulo utilitário para renomeamento de arquivos de relatórios contábeis
"""
import os
import re
from datetime import datetime
from typing import Optional

def rename_report_file(
    current_filename: str,
    codi_emp: str,
    data_inicial: str,
    data_final: str,
    timestamp_str: str,
    idioma_ingles: bool = False,
    tipo_relatorio: Optional[str] = None,
    conn_str: str = None
) -> str:
    """
    Renomeia um arquivo de relatório conforme o padrão especificado

    Args:
        current_filename: Nome atual do arquivo
        codi_emp: Código da empresa
        data_inicial: Data inicial no formato YYYY-MM-DD
        data_final: Data final no formato YYYY-MM-DD
        timestamp_str: Timestamp da geração
        idioma_ingles: Se True, usa nomes em inglês
        tipo_relatorio: Tipo do relatório (opcional, será inferido do nome se não informado)

    Returns:
        Novo nome do arquivo
    """


    # Determinar o tipo de relatório baseado no nome do arquivo se não foi informado
    if tipo_relatorio is None:
        # Remove sufixos de idioma (_EN) para melhor detecção
        filename_clean = current_filename.replace('_EN', '')
        if 'Balancete' in filename_clean or 'balancete' in filename_clean:
            tipo_relatorio = 'balancete'
        elif 'Comparativo' in filename_clean or 'comparativo' in filename_clean:
            tipo_relatorio = 'comparativo'
        elif 'DRE' in filename_clean or 'dre' in filename_clean:
            tipo_relatorio = 'dre'
        elif 'razao' in filename_clean or 'Razao' in filename_clean:
            tipo_relatorio = 'razao'
        else:
            # Se não conseguir determinar, retorna o nome original
            print(f"Debug: Tipo de relatório não identificado no nome: {current_filename}")
            return current_filename

    print(f"Debug: Tipo de relatório identificado: {tipo_relatorio}")

    # Buscar nome da empresa
    if conn_str is not None:
        nome_empresa = buscar_nome_empresa(codi_emp, conn_str)
        # Remove o código do início do nome, se houver
        nome_empresa = nome_empresa.strip()
        if nome_empresa.startswith(str(codi_emp)):
            nome_empresa = nome_empresa[len(str(codi_emp)):].lstrip(' -_')
    else:
        nome_empresa = str(codi_emp)
    print(f"Debug: Nome da empresa: {nome_empresa}")

    # Extrair extensão do arquivo
    _, ext = os.path.splitext(current_filename)

    # Determinar se é acumulado ou mensal baseado nas datas
    data_inicial_dt = datetime.strptime(data_inicial, '%Y-%m-%d')
    data_final_dt = datetime.strptime(data_final, '%Y-%m-%d')

    # Se o período for do início do ano até o mês final, é acumulado
    # Caso contrário, é mensal
    is_acumulado = (
        data_inicial_dt.month == 1 and
        data_inicial_dt.day == 1
    )

    print(f"Debug: Período {data_inicial} a {data_final}, is_acumulado = {is_acumulado}")

    # Formatar mês e ano da competência
    mes_competencia = data_final_dt.strftime('%m')
    ano_competencia = data_final_dt.strftime('%Y')
    competencia_formatada = f"{mes_competencia}.{ano_competencia}"

    # Formatar data de geração
    data_geracao = datetime.now()
    dia_geracao = data_geracao.strftime('%d')
    mes_geracao = data_geracao.strftime('%m')
    ano_geracao = data_geracao.strftime('%Y')
    emissao_formatada = f"{dia_geracao}.{mes_geracao}.{ano_geracao}"

    # Construir novo nome baseado no tipo e idioma

    if idioma_ingles:
        if tipo_relatorio == 'balancete':
            tipo_nome = 'Trial Balance Sheet Ytd' if is_acumulado else 'Trial Balance Sheet Monthly'
        elif tipo_relatorio == 'comparativo':
            tipo_nome = 'Result Movement Ytd'
        elif tipo_relatorio == 'dre':
            tipo_nome = 'P&L Ytd' if is_acumulado else 'P&L Monthly'
        elif tipo_relatorio == 'razao':
            tipo_nome = 'General Ledger Monthly'
        else:
            tipo_nome = tipo_relatorio

        novo_nome = f"{nome_empresa} - {tipo_nome} {competencia_formatada} - Issued on {emissao_formatada}{ext}"

    else:  # Português
        if tipo_relatorio == 'balancete':
            tipo_nome = 'Balancete Acumulado' if is_acumulado else 'Balancete Mensal'
        elif tipo_relatorio == 'comparativo':
            tipo_nome = 'Comparativo de Resultado Acumulado'
        elif tipo_relatorio == 'dre':
            tipo_nome = 'D.R.E. Acumulado' if is_acumulado else 'D.R.E. Mensal'
        elif tipo_relatorio == 'razao':
            tipo_nome = 'Razão Mensal'
        else:
            tipo_nome = tipo_relatorio

        novo_nome = f"{nome_empresa} - {tipo_nome} {competencia_formatada} - Emitido em {emissao_formatada}{ext}"

    return novo_nome

def rename_file_after_generation(
    current_path: str,
    codi_emp: str,
    data_inicial: str,
    data_final: str,
    timestamp_str: str,
    idioma_ingles: bool = False,
    tipo_relatorio: Optional[str] = None,
    conn_str: str = None
) -> str:
    """
    Renomeia um arquivo após sua geração

    Args:
        current_path: Caminho completo do arquivo atual
        codi_emp: Código da empresa
        data_inicial: Data inicial no formato YYYY-MM-DD
        data_final: Data final no formato YYYY-MM-DD
        timestamp_str: Timestamp da geração
        idioma_ingles: Se True, usa nomes em inglês
        tipo_relatorio: Tipo do relatório

    Returns:
        Novo caminho do arquivo
    """

    if not os.path.exists(current_path):
        print(f"Aviso: Arquivo {current_path} não encontrado para renomeamento")
        return current_path

    current_dir = os.path.dirname(current_path)
    current_filename = os.path.basename(current_path)

    print(f"Debug: Iniciando renomeamento de {current_filename}")
    print(f"Debug: Caminho completo: {current_path}")
    print(f"Debug: Arquivo existe: {os.path.exists(current_path)}")
    print(f"Debug: Parâmetros - codi_emp: {codi_emp}, data_inicial: {data_inicial}, data_final: {data_final}, idioma_ingles: {idioma_ingles}")

    novo_filename = rename_report_file(
        current_filename, codi_emp, data_inicial, data_final,
        timestamp_str, idioma_ingles, tipo_relatorio, conn_str
    )

    print(f"Debug: Novo nome gerado: {novo_filename}")

    novo_path = os.path.join(current_dir, novo_filename)
    print(f"Debug: Novo caminho: {novo_path}")

    # Renomear o arquivo
    try:
        os.rename(current_path, novo_path)
        print(f"Arquivo renomeado: {current_filename} -> {novo_filename}")
        # Extrair ano e mês de competência do nome (já formatado em MM.YYYY)
        partes = novo_filename.split()
        mes_ano = None
        for p in partes:
            if re.match(r"\d{2}\.\d{4}", p):
                mes_ano = p
                break
        if mes_ano:
            mes, ano = mes_ano.split('.')
            mover_arquivo_para_destino(novo_path, codi_emp, ano, mes)
        else:
            print("Não foi possível extrair mês e ano de competência do nome do arquivo.")
        return novo_path
    except Exception as e:
        print(f"Erro ao renomear arquivo {current_filename}: {e}")
        return current_path
