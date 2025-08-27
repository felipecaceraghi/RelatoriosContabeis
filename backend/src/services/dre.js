const { runPython } = require('./pythonHelper');

async function run({ codiEmp, dataInicial, dataFinal, ingles }) {
  const moduleName = 'DRE';
  const funcName = 'gerar_dre';
  const payload = { codi_emp: codiEmp, data_inicial: dataInicial, data_final: dataFinal, ingles };
  const res = await runPython(moduleName, funcName, payload);
  return res;
}

module.exports = { run };
