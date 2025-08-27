const express = require('express');
const router = express.Router();
const { getCompanies } = require('../services/companies');

router.get('/', async (req, res) => {
  const q = req.query.q || req.query.qterm || null;
  try{
    const rows = await getCompanies(q);
    res.json(rows);
  }catch(err){
    console.error('Error fetching companies', err);
    res.status(500).json({ error: 'Error fetching companies', detail: err.message });
  }
});

module.exports = router;
