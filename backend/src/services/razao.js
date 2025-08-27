const { runPython } = require('./pythonHelper');

async function run({ codiEmp, dataInicial, dataFinal, ingles }) {
  // Map to Python script RAZAO.py -> gerar_relatorio_razao_com_dump
  const moduleName = 'RAZAO';
  const funcName = 'gerar_relatorio_razao_com_dump';
  const payload = { codi_emp: codiEmp, data_inicial: dataInicial, data_final: dataFinal, filiais: false, idioma_ingles: ingles };
  const res = await runPython(moduleName, funcName, payload);
  return res;
}

module.exports = { run };
