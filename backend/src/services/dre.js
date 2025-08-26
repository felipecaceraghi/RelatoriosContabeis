async function run({ codiEmp, dataInicial, dataFinal, ingles }) {
  return Promise.resolve({ message: 'dre service called', codiEmp, dataInicial, dataFinal, ingles });
}

module.exports = { run };
