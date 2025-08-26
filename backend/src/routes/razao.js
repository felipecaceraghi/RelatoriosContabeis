const express = require('express');
const router = express.Router();
const razaoService = require('../services/razao');

router.post('/', async (req, res) => {
  const { codiEmp, dataInicial, dataFinal, ingles } = req.body;

  // basic validation
  if (!Number.isInteger(codiEmp) || typeof dataInicial !== 'string' || typeof dataFinal !== 'string' || typeof ingles !== 'boolean') {
    return res.status(400).json({ error: 'Invalid parameters. Expected codiEmp:int, dataInicial:string, dataFinal:string, ingles:boolean' });
  }

  try {
    const result = await razaoService.run({ codiEmp, dataInicial, dataFinal, ingles });
    res.json({ success: true, data: result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
