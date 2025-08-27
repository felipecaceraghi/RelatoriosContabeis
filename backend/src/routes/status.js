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

  res.json({ success: true, status });
});

module.exports = router;
