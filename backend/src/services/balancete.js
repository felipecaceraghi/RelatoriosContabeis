const { runPython } = require('./pythonHelper');

async function run({ codiEmp, dataInicial, dataFinal, ingles }) {
  const moduleName = 'BALANCETE';
  const funcName = 'GerarRelatorioBalancete';
  const payload = { codi_emp: codiEmp, data_inicial: dataInicial, data_final: dataFinal, ingles };
  const res = await runPython(moduleName, funcName, payload);
  return res;
}

module.exports = { run };
