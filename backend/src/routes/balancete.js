const express = require('express');
const router = express.Router();
const balanceteService = require('../services/balancete');

router.post('/', async (req, res) => {
  const { codiEmp, dataInicial, dataFinal, ingles } = req.body;
  if (!Number.isInteger(codiEmp) || typeof dataInicial !== 'string' || typeof dataFinal !== 'string' || typeof ingles !== 'boolean') {
    return res.status(400).json({ error: 'Invalid parameters. Expected codiEmp:int, dataInicial:string, dataFinal:string, ingles:boolean' });
  }

  try {
    const result = await balanceteService.run({ codiEmp, dataInicial, dataFinal, ingles });
    res.json({ success: true, data: result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
