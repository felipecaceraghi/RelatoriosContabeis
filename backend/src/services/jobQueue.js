const { randomUUID } = require('crypto');
const fs = require('fs');
const path = require('path');

// simple in-memory job queue
const queue = [];
const jobs = new Map();
let running = 0;
const CONCURRENCY = 2;

// directory where we will save raw python outputs for debugging
const OUTPUT_DIR = path.resolve(__dirname, '../../output');
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// helper: persist job JSON to disk so status survives restarts
function saveJobToDisk(job){
  try{
    const safe = Object.assign({}, job);
    // remove anything non-serializable
    if(safe._internal) delete safe._internal;
    const p = path.join(OUTPUT_DIR, `${job.id}.json`);
    fs.writeFileSync(p, JSON.stringify(safe, null, 2), { encoding: 'utf8' });
  }catch(e){ /* ignore */ }
}

// on startup, load persisted jobs so we can inspect previous runs
function loadPersistedJobs(){
  try{
    const files = fs.readdirSync(OUTPUT_DIR);
    files.forEach(f => {
      if(f.endsWith('.json')){
        try{
          const content = fs.readFileSync(path.join(OUTPUT_DIR, f), 'utf8');
          const job = JSON.parse(content);
          if(job && job.id) jobs.set(job.id, job);
        }catch(e){ /* ignore parse errors */ }
      }
    });
  }catch(e){ /* ignore */ }
}

loadPersistedJobs();

// map job type to service module
const serviceMap = {
  razao: require('./razao'),
  dre: require('./dre'),
  comparativo: require('./comparativo'),
  balancete: require('./balancete')
};

function enqueue(type, payload){
  const id = randomUUID();
  const job = { id, type, payload, status: 'pending', progress: 0, result: null, error: null, createdAt: Date.now() };
  jobs.set(id, job);
  queue.push(job);
  // persist immediately
  saveJobToDisk(job);
  processQueue();
  return id;
}

function getStatus(id){
  return jobs.get(id) || null;
}

async function processQueue(){
  while(running < CONCURRENCY && queue.length > 0){
    const job = queue.shift();
    if(!job) break;
    running++;
    job.status = 'running';
    job.startedAt = Date.now();
    // run async but do not block loop
    (async ()=>{
      try{
        const svc = serviceMap[job.type];
        if(!svc || typeof svc.run !== 'function') throw new Error('Unknown service for job type: '+job.type);
        // allow service to run; it may be long
        const res = await svc.run(job.payload);
  job.status = 'complete';
  job.progress = 100;
  job.result = res;
  // persist job metadata
  saveJobToDisk(job);
        // if python returned raw text or other data, persist it for inspection
        try{
          const content = (typeof res === 'string' || Buffer.isBuffer(res)) ? String(res) : JSON.stringify(res, null, 2);
          const outPath = path.join(OUTPUT_DIR, `${job.id}.log`);
          fs.writeFileSync(outPath, content, { encoding: 'utf8' });
          job.logFile = outPath;
        }catch(e){ /* non-fatal */ }
        // if result is a list of filenames, try to copy them from scripts dir to OUTPUT_DIR
        try{
          if(Array.isArray(res) && res.length > 0){
            const scriptsDir = path.resolve(__dirname, '../../scripts');
            const copied = [];
            for(const f of res){
              try{
                const src = path.join(scriptsDir, f);
                const dest = path.join(OUTPUT_DIR, f);
                if(fs.existsSync(src)){
                  fs.copyFileSync(src, dest);
                  copied.push(dest);
                }
              }catch(e){/* ignore individual file copy errors */}
            }
            if(copied.length) {
              job.copiedFiles = copied;
              saveJobToDisk(job);
            }
          }
        }catch(e){/* ignore */}
        job.completedAt = Date.now();
      }catch(err){
  job.status = 'failed';
  job.error = (err && err.message) ? err.message : String(err);
  job.completedAt = Date.now();
  // persist job metadata
  saveJobToDisk(job);
        // save error details to disk
        try{
          const outPath = path.join(OUTPUT_DIR, `${job.id}.log`);
          const content = (err && err.stack) ? err.stack : job.error;
          fs.writeFileSync(outPath, String(content), { encoding: 'utf8' });
          job.logFile = outPath;
        }catch(e){ /* ignore */ }
      }finally{
        running--;
        // process next in queue
        setImmediate(processQueue);
      }
    })();
  }
}

module.exports = { enqueue, getStatus };
