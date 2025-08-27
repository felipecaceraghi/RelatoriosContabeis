# Relatorios Contabeis - Backend

Projeto Node.js + Express minimal para as rotas: `razao`, `dre`, `comparativo`, `balancete`.

Endpoints (POST) - todas recebem JSON:

- /razao
- /dre
- /comparativo
- /balancete

Payload esperado em todas as rotas:

{
  "codiEmp": 1,
  "dataInicial": "2025-01-01",
  "dataFinal": "2025-12-31",
  "ingles": false
}

Instalação

1. Abra um terminal em `backend`.
2. Rode `npm install`.
3. Rode `npm start`.

Python integration
------------------
O backend pode chamar funções Python presentes na pasta `scripts` do workspace.

- Local das scripts padrão: `C:\Users\estatistica007\Documents\nextProjects\RelatoriosContabeis\scripts`.
- Você pode sobrescrever com a variável de ambiente `PY_SCRIPTS_DIR`.
- É necessário ter Python 3 instalado e disponível no PATH como `python`.
- O runner (`src/services/py_runner.py`) espera JSON via stdin e retorna JSON via stdout.

Recomendações:
- Garanta que os arquivos `RAZAO.py`, `DRE.py`, `COMPARATIVO.py`, `BALANCETE.py` exportem as funções mencionadas na especificação.
- Para tarefas longas, considere um sistema de fila em vez de executar síncronamente.
