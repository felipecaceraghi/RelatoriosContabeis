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
