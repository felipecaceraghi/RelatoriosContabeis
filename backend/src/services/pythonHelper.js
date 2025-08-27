const { spawn } = require('child_process');
const path = require('path');

const PY_RUNNER = path.join(__dirname, 'py_runner.py');

function tryPythonExecutables(){
  return ['python', 'python3'];
}

function runPython(moduleName, functionName, payload = {}){
  return new Promise((resolve, reject) => {
    const candidates = tryPythonExecutables();
    let tried = 0;
    let lastErr = null;

    const attempt = (exeIndex) => {
      const exe = candidates[exeIndex];
        // ensure python child uses UTF-8 for stdin/stdout to avoid encoding errors on Windows
        // ensure python child uses UTF-8 for stdin/stdout and knows where scripts and outputs live
        const childEnv = Object.assign({}, process.env, {
          PYTHONIOENCODING: 'utf-8',
          PYTHONUTF8: '1',
          PY_SCRIPTS_DIR: path.resolve(__dirname, '../../../scripts'),
          PY_OUTPUT_DIR: path.resolve(__dirname, '../../output')
        });
        const py = spawn(exe, [PY_RUNNER, moduleName, functionName], { env: childEnv });

      let out = '';
      let err = '';
      py.stdout.on('data', d => out += d.toString());
      py.stderr.on('data', d => err += d.toString());

      py.on('error', e => {
        lastErr = e;
        tried++;
        if(tried < candidates.length) return attempt(tried);
        return reject(lastErr || e);
      });

      py.on('close', code => {
        if(err) console.error('py stderr:', err);
        if(!out) return reject(new Error('No output from python runner'));
        try{
          const j = JSON.parse(out);
          if(code !== 0) return reject(new Error(j.error || ('python exited ' + code)));
          resolve(j);
        }catch(e){
          resolve({ raw: out, code });
        }
      });

      try{ py.stdin.write(JSON.stringify(payload)); }catch(e){}
      try{ py.stdin.end(); }catch(e){}
    };

    attempt(0);
  });
}

module.exports = { runPython };
