const { runPython } = require('./pythonHelper');

async function run({ codiEmp, dataInicial, dataFinal, ingles }) {
  const moduleName = 'COMPARATIVO';
  const funcName = 'gerar_relatorio';
  const payload = { codigo_empresa: codiEmp, data_inicio: dataInicial, data_fim: dataFinal, ingles };
  const res = await runPython(moduleName, funcName, payload);
  return res;
}

module.exports = { run };
