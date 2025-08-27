const express = require('express');
const router = express.Router();
const dreService = require('../services/dre');
const { enqueue } = require('../services/jobQueue');

router.post('/', async (req, res) => {
  const { codiEmp, dataInicial, dataFinal, ingles } = req.body;
  if (!Number.isInteger(codiEmp) || typeof dataInicial !== 'string' || typeof dataFinal !== 'string' || typeof ingles !== 'boolean') {
    return res.status(400).json({ error: 'Invalid parameters. Expected codiEmp:int, dataInicial:string, dataFinal:string, ingles:boolean' });
  }

  try {
    const processing_id = enqueue('dre', { codiEmp, dataInicial, dataFinal, ingles });
    res.json({ success: true, processing_id });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
