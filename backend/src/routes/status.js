const express = require('express');
const router = express.Router();
const { getStatus } = require('../services/jobQueue');
const fs = require('fs');
const path = require('path');

// same OUTPUT_DIR as jobQueue
const OUTPUT_DIR = path.resolve(__dirname, '../../output');
if (!fs.existsSync(OUTPUT_DIR)) {
  try{ fs.mkdirSync(OUTPUT_DIR, { recursive: true }); }catch(e){}
}

router.get('/:id', (req, res) => {
  const id = req.params.id;
  const status = getStatus(id);
  if(!status) return res.status(404).json({ error: 'Job not found' });

  // If job completed/failed but no logFile exists (older job before log persistence), create one from result/error
  if(!status.logFile && (status.status === 'complete' || status.status === 'failed')){
    try{
      const outPath = path.join(OUTPUT_DIR, `${id}.log`);
      let content = '';
      if(status.result) content += typeof status.result === 'string' ? status.result : JSON.stringify(status.result, null, 2);
      if(status.error) content += '\n\nERROR:\n' + String(status.error);
      if(content){
        fs.writeFileSync(outPath, content, { encoding: 'utf8' });
        status.logFile = outPath;
      }
    }catch(e){ /* ignore */ }
  }

  // Adiciona campo files no result, se possível
  let files = [];
  if (status.result && typeof status.result === 'object') {
    // Se for dict com pdf/xlsx/json
    if (status.result.pdf || status.result.xlsx || status.result.json) {
      if (status.result.pdf) files.push({ name: path.basename(status.result.pdf), path: status.result.pdf });
      if (status.result.xlsx) files.push({ name: path.basename(status.result.xlsx), path: status.result.xlsx });
      if (status.result.json) files.push({ name: path.basename(status.result.json), path: status.result.json });
    }
    // Se for lista de arquivos
    if (Array.isArray(status.result)) {
      status.result.forEach(f => files.push({ name: path.basename(f), path: f }));
    }
  }
  // Se copiou arquivos
  if (status.copiedFiles && Array.isArray(status.copiedFiles)) {
    status.copiedFiles.forEach(f => files.push({ name: path.basename(f), path: f }));
  }
  // Remove duplicados
  files = files.filter((v,i,a)=>a.findIndex(t=>(t.path===v.path))===i);
  // Adiciona no result
  if (!status.result) status.result = {};
  status.result.files = files;

  res.json({ success: true, status });
});

module.exports = router;
