import { useState, useEffect, useRef } from 'react'
import Header from '../components/Header'
import Sidebar from '../components/Sidebar'
import Calendar from '../components/Calendar'

export default function Home(){
  // base URL for backend API (set NEXT_PUBLIC_API_URL in .env.local if backend runs on a different port)
  const API_BASE = process.env.NEXT_PUBLIC_API_URL || ''
  const [companies, setCompanies] = useState([])
  const [query, setQuery] = useState('')
  const [dropdownOpen, setDropdownOpen] = useState(false)
  const [selectedCompanies, setSelectedCompanies] = useState([])
  const [dataInicio, setDataInicio] = useState('')
  const [dataFim, setDataFim] = useState('')
  const [relatorios, setRelatorios] = useState({ balancete: true, comparativo: true, dre: true, razao: true })
  const [idiomas, setIdiomas] = useState({ portugues: true, ingles: false })
  
  const [processing, setProcessing] = useState(false)
  const [progress, setProgress] = useState(0)
  const [progressMessage, setProgressMessage] = useState('')
  const [modal, setModal] = useState({ open: false, success: false, message: '' })
  const dropdownRef = useRef(null)
  const fetchTimer = useRef(null)

  useEffect(()=>{ // set default dates to current month
    const today = new Date();
    const firstDay = new Date(today.getFullYear(), today.getMonth(), 1).toISOString().slice(0,10)
    const lastDay = new Date(today.getFullYear(), today.getMonth()+1, 0).toISOString().slice(0,10)
    setDataInicio(firstDay); setDataFim(lastDay)
  },[])

  useEffect(()=>{ // initial load - fetch a small set or all
    fetchCompanies('')
  },[])

  useEffect(()=>{ const onClick = (e)=>{ if(dropdownRef.current && !dropdownRef.current.contains(e.target)) setDropdownOpen(false)}; document.addEventListener('click', onClick); return ()=>document.removeEventListener('click', onClick)},[])

  // filtered list comes from server; we still require min 2 chars to show suggestions
  const filtered = query.length >= 2 ? companies : []

  const addCompany = (c)=>{ if(!selectedCompanies.some(s=>s.code===c.code)){ setSelectedCompanies(s=>[...s,c]) } setQuery(''); setDropdownOpen(false) }
  const removeCompany = (code)=> setSelectedCompanies(s=>s.filter(x=>x.code!==code))

  // fetch companies from server with optional search term
  const fetchCompanies = async (q)=>{
    try{
      const path = q && q.length>=2 ? `/companies?q=${encodeURIComponent(q)}` : '/companies'
      const { res, ok } = await safeFetch(path, { method: 'GET' })
      if(!res || !ok) {
        console.debug('fetchCompanies: no results or endpoint unreachable', { path, apiBase: API_BASE })
        return setCompanies([])
      }
      const json = await res.json()
      console.debug('fetchCompanies: got', json?.length || 0, 'results for', q)
      setCompanies(json)
    }catch(e){ setCompanies([]) }
  }

  // Helper to call backend endpoints safely (tries API_BASE -> relative -> localhost:3000)
  const safeFetch = async (path, options={})=>{
    const tried = []
    const build = (base)=> base ? `${base.replace(/\/$/, '')}${path}` : path
    // candidates: API_BASE (if set), relative, localhost fallback
    const candidates = []
    if(API_BASE) candidates.push(API_BASE)
    candidates.push('')
    candidates.push('https://contabilestatistica.loca.lt')

    // Always add bypass-tunnel-reminder header for localtunnel compatibility
    const mergedOptions = {
      ...options,
      headers: {
        ...(options.headers || {}),
        'bypass-tunnel-reminder': 'true',
      },
    }
    // Adiciona access-control-request-headers manualmente para forçar preflight com bypass-tunnel-reminder
    if (!mergedOptions.headers['access-control-request-headers']) {
      mergedOptions.headers['access-control-request-headers'] = 'bypass-tunnel-reminder,content-type';
    }

    for(const base of candidates){
      const url = build(base)
      tried.push(url)
      try{
        const res = await fetch(url, mergedOptions)
        // if not ok, but it's HTML from Next, try next
        const contentType = res.headers.get('content-type') || ''
        if(!res.ok || contentType.includes('text/html')){
          // if last candidate, return response for error handling
          if(base === candidates[candidates.length-1]) return { res, ok: res.ok }
          continue
        }
        // good response and not html
        return { res, ok: true }
      }catch(e){
        // network error - try next
        if(base === candidates[candidates.length-1]) return { res: null, ok:false, error: e }
        continue
      }
    }
    return { res: null, ok:false }
  }

  // debounce query -> server
  useEffect(()=>{
    if(fetchTimer.current) clearTimeout(fetchTimer.current)
    // only search when 2+ chars
    fetchTimer.current = setTimeout(()=>{
      fetchCompanies(query)
    }, 300)
    return ()=>{ if(fetchTimer.current) clearTimeout(fetchTimer.current) }
  },[query])

  const toggleRel = (k)=> setRelatorios(r=>({...r, [k]: !r[k]}))
  const toggleIdioma = (k)=> setIdiomas(i=>({...i, [k]: !i[k]}))

  const handleSubmit = async (e)=>{
    e.preventDefault()
    if(selectedCompanies.length===0) return alert('Selecione pelo menos uma empresa')
    if(!dataInicio||!dataFim) return alert('Preencha as datas')
    // build jobs: one per selected company x selected report x idioma x (normal/acumulado)
    const selectedReports = Object.keys(relatorios).filter(k=>relatorios[k])
    if(selectedReports.length===0) return alert('Selecione ao menos um relatório')

    const jobs = []
    try{
      setProcessing(true); setProgress(0); setProgressMessage('Enfileirando trabalhos...')
      for(const comp of selectedCompanies){
        // determine which language calls to make
        const languageSelections = []
        if(idiomas.portugues && idiomas.ingles){
          languageSelections.push(false, true)
        }else if(idiomas.portugues){
          languageSelections.push(false)
        }else if(idiomas.ingles){
          languageSelections.push(true)
        }else{
          languageSelections.push(false)
        }

        for(const rpt of selectedReports){
          const route = `/${rpt}`
          const isAcumulado = (rpt==='balancete'||rpt==='dre')
          for(const lang of languageSelections){
            // sempre gera para o período escolhido
            const bodyNormal = { codiEmp: Number(comp.code), dataInicial: dataInicio, dataFinal: dataFim, ingles: lang }
            const { res: pres, ok } = await safeFetch(route, { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(bodyNormal) })
            if(!ok){
              const text = pres ? await pres.text().catch(()=>'<no body>') : 'no response'
              console.error('Failed to enqueue', route, pres, text)
              continue
            }
            let j
            try{ j = await pres.json() }catch(e){ const text = await pres.text().catch(()=>'<no body>'); console.error('Invalid JSON from enqueue', text); continue }
            if(j.success && j.processing_id){ jobs.push(j.processing_id) } else { console.error('Failed to enqueue', route, j); }

            // se for balancete ou dre, também gera acumulado
            if(isAcumulado){
              // data inicial acumulada: 01-01-ANO da dataInicio
              const ano = dataInicio.slice(0,4)
              const dataIniAcum = `${ano}-01-01`
              // só gera acumulado se o período não for já 01-01-ANO até dataFim
              if(dataInicio !== dataIniAcum){
                const bodyAcum = { codiEmp: Number(comp.code), dataInicial: dataIniAcum, dataFinal: dataFim, ingles: lang }
                const { res: presA, ok: okA } = await safeFetch(route, { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(bodyAcum) })
                if(!okA){
                  const textA = presA ? await presA.text().catch(()=>'<no body>') : 'no response'
                  console.error('Failed to enqueue acumulado', route, presA, textA)
                  continue
                }
                let jA
                try{ jA = await presA.json() }catch(e){ const textA = await presA.text().catch(()=>'<no body>'); console.error('Invalid JSON from enqueue acumulado', textA); continue }
                if(jA.success && jA.processing_id){ jobs.push(jA.processing_id) } else { console.error('Failed to enqueue acumulado', route, jA); }
              }
            }
          }
        }
      }

      if(jobs.length===0){ setProcessing(false); return alert('Nenhum trabalho foi enfileirado (verifique o console)') }

      setProgressMessage(`Aguardando conclusão de ${jobs.length} trabalho(s)...`)

      // poll all jobs until complete
      const statuses = new Map()
      let remaining = jobs.length
      const pollInterval = 1500
      await new Promise((resolve)=>{
        const iv = setInterval(async ()=>{
          for(const id of jobs){
            if(statuses.get(id)?.status === 'complete' || statuses.get(id)?.status === 'failed') continue
            try{
              const statusPath = `/status/${id}`
              const { res: sres, ok: sok } = await safeFetch(statusPath)
              if(!sok){ statuses.set(id, { status: 'error' }); continue }
              let sj
              try{ sj = await sres.json() }catch(e){ const t = await sres.text().catch(()=>'<no body>'); console.error('Invalid status json', t); statuses.set(id, { status: 'error' }); continue }
              const st = sj.status || sj
              statuses.set(id, st)
            }catch(e){ statuses.set(id, { status: 'error' }) }
          }

          // compute progress
          let done = 0
          let failed = 0
          for(const [id, st] of statuses.entries()){
            if(!st) continue
            if(st.status === 'complete') done++
            if(st.status === 'failed') failed++
          }
          remaining = jobs.length - done - failed
          const pct = Math.round((done / jobs.length) * 100)
          setProgress(pct)
          setProgressMessage(`Concluídos: ${done}, Falhas: ${failed}, Aguardando: ${Math.max(0, remaining)}`)

          if((done + failed) >= jobs.length){ clearInterval(iv); resolve(); }
        }, pollInterval)
      })

      setProcessing(false)
      // show result summary
      // Busca detalhes dos relatórios gerados (nomes e caminhos) se possível
      let detalhes = []
      for(const [id, st] of statuses.entries()){
        if(st && st.status === 'complete' && st.result && st.result.files){
          // Aceita tanto objetos quanto strings em result.files
          st.result.files.forEach(f => {
            let nome, caminho
            if(typeof f === 'string') {
              caminho = f
              // Extrai apenas o nome do arquivo do caminho
              nome = f.split(/[\\/]/).pop()
            } else if(typeof f === 'object' && f !== null) {
              nome = f.name || (f.path ? f.path.split(/[\\/]/).pop() : '') || ''
              caminho = f.path || f.name || ''
            } else {
              nome = String(f)
              caminho = String(f)
            }
            detalhes.push(`✔️ ${nome} gerado em: ${caminho}`)
          })
        } else if(st && st.status === 'complete') {
          detalhes.push(`✔️ Relatório gerado com sucesso (ID: ${id})`)
        } else if(st && st.status === 'failed') {
          detalhes.push(`❌ Erro ao gerar relatório (ID: ${id})`)
        }
      }
      if(detalhes.length === 0) detalhes.push('Todos os trabalhos finalizados, mas não foi possível obter detalhes dos arquivos gerados.')
      setModal({ open: true, success: true, message: detalhes.join('\n') })

    }catch(err){ setProcessing(false); alert('Erro: '+err.message) }
  }

  // When a start date is chosen, set the end date to the last day of that month
  const handleStartDateChange = (isoDate)=>{
    setDataInicio(isoDate)
    if(!isoDate) return
    const parts = isoDate.split('-')
    if(parts.length!==3) return
    const y = Number(parts[0]), m = Number(parts[1])
    const last = new Date(y, m, 0) // month is 1-based in iso; new Date(y,m,0) is last day of month m
    const lastIso = `${last.getFullYear()}-${String(last.getMonth()+1).padStart(2,'0')}-${String(last.getDate()).padStart(2,'0')}`
    setDataFim(lastIso)
  // (no acumulados behavior — acumulados removed)
  }

  

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-gradient-to-r from-sky-700 to-sky-900 text-white py-3 shadow-lg">
        <div className="container mx-auto px-4 flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-4">
          <img 
            src="/logo.png" 
            alt="Logo" 
            className="h-[43.2px] sm:h-12 object-contain mb-2 sm:mb-0 mx-auto sm:mx-0 transition-all duration-200"
            style={{ height: '43.2px' }} // fallback for mobile (10% menor que 48px)
          />
          <h1 
            className="text-xl font-bold whitespace-normal text-center sm:text-left w-full sm:w-auto"
            style={{ textAlign: 'center' }}
          >
            <span className="block sm:inline">Sistema de Automação - Relatórios Contábeis</span>
          </h1>
        </div>
        <style jsx>{`
          @media (min-width: 640px) {
            header .container {
              justify-content: flex-start !important;
            }
            header img {
              margin-left: 0 !important;
              margin-right: 0 !important;
            }
            header h1 {
              text-align: left !important;
            }
          }
          @media (max-width: 639px) {
            header .container {
              justify-content: center !important;
            }
            header img {
              height: 43.2px !important;
              margin-left: auto !important;
              margin-right: auto !important;
            }
            header h1 {
              text-align: center !important;
            }
          }
        `}</style>
      </header>

      <div className="app-scale">
      <div className="container mx-auto px-2 sm:px-4 py-4 sm:py-6">
        <div className={`mb-6 ${processing? 'block':'hidden'}`}>
          <div className="card-panel card-notice">
            <div className="flex items-center gap-3">
              <div className={`w-10 h-10 rounded-full flex items-center justify-center text-white ${processing? 'bg-sky-400':'bg-green-500'}`}><i className="fas fa-database"></i></div>
              <div className="flex-1">
                <div className="font-medium">{processing? 'Processamento em andamento' : 'Banco de Dados Conectado'}</div>
                <div className="text-sm text-muted">{progressMessage||'Verificando status...'}</div>
              </div>
            </div>
          </div>
        </div>

        <div className="card card-panel">
          <form onSubmit={handleSubmit} className="card-body">
            {/* Top: Basic info */}
            <div className="mb-6">
              <h3 className="text-lg font-semibold mb-3 flex items-center gap-2"><i className="fas fa-info-circle text-sky-600"></i> Informações Básicas</h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <label className="text-sm font-medium">Empresas</label>
                    <div className="text-xs bg-sky-400 text-white px-2 py-1 rounded">{selectedCompanies.length} selecionada{selectedCompanies.length!==1?'s':''}</div>
                  </div>
                  <div className="relative" ref={dropdownRef}>
                    <input value={query} onChange={e=>{setQuery(e.target.value); setDropdownOpen(true)}} onFocus={()=>setDropdownOpen(true)} placeholder="Digite para buscar e selecionar..." className="w-full border-2 border-gray-200 rounded p-3 h-12" />
                    {dropdownOpen && filtered.length>0 && (
                      <div className="absolute z-30 bg-white border rounded mt-2 w-full max-h-56 overflow-auto shadow-md">
                        {filtered.map(c=> (
                          <div key={c.code} className="p-3 border-b last:border-b-0 hover:bg-sky-100 cursor-pointer" onClick={()=>addCompany(c)}>
                            <div className="font-medium">{c.name}</div>
                            <div className="text-xs text-gray-500">Código: {c.code}</div>
                          </div>
                        ))}
                      </div>
                    )}
                    <div className="flex flex-wrap gap-2 mt-3">
                      {selectedCompanies.map(c=> (
                        <div key={c.code} className="inline-flex items-center gap-2 bg-gradient-to-r from-sky-400 to-sky-600 text-white px-3 py-1 rounded-full shadow-sm">
                          <span className="text-sm">{c.name} ({c.code})</span>
                          <button type="button" onClick={()=>removeCompany(c.code)} className="bg-white/20 rounded-full w-5 h-5 flex items-center justify-center text-xs">×</button>
                        </div>
                      ))}
                    </div>
                    {dropdownOpen && query.length>=2 && filtered.length===0 && (
                      <div className="absolute z-30 bg-white border rounded mt-2 w-full p-3 text-sm text-gray-600">Nenhum resultado encontrado.</div>
                    )}
                  </div>
                </div>

                <div>
                  <label className="text-sm font-medium">Período de Competência (Mensal)</label>
                  <div className="grid grid-cols-2 gap-4 mt-2">
                    <Calendar value={dataInicio} onChange={d=>handleStartDateChange(d)} placeholder="Data inicial" />
                    <Calendar value={dataFim} onChange={d=>setDataFim(d)} placeholder="Data final" />
                  </div>
                </div>
              </div>
            </div>

            {/* Middle: Relatórios | Idioma */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 items-stretch">
              <div className="group-box">
                <div className="group-header"><i className="fas fa-chart-line text-sky-600 mr-2"></i>Relatórios Contábeis</div>
                <div className="group-body">
                  <div className="grid grid-cols-2 gap-4 sm:gap-6">
                    {['balancete','comparativo','dre','razao'].map(k=> (
                      <div key={k} onClick={()=>toggleRel(k)} className={`tile ${relatorios[k]? 'tile-selected':'tile-unselected'}`}>
                        <div className="font-medium text-sm capitalize">{k.replace('_',' ')}</div>
                        <div className="check">✓</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div className="group-box">
                <div className="group-header"><i className="fas fa-language text-sky-600 mr-2"></i>Idioma dos Relatórios</div>
                <div className="group-body">
                  <div className="grid grid-cols-1 gap-4 sm:gap-6">
                    <div onClick={()=>toggleIdioma('portugues')} className={`tile ${idiomas.portugues? 'tile-selected':'tile-unselected'}`}>
                      <div className="font-medium text-sm">Português</div>
                      <div className="check">✓</div>
                    </div>
                    <div onClick={()=>toggleIdioma('ingles')} className={`tile ${idiomas.ingles? 'tile-selected':'tile-unselected'}`}>
                      <div className="font-medium text-sm">Inglês</div>
                      <div className="check">✓</div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Opções Adicionais removidas intentionally */}
            </div>
            {/* Actions - button below everything */}
            <div className="mt-6">
              <div className="grid grid-cols-1 gap-3">
                <button type="submit" className="btn-primary w-full">Gerar Relatórios Contábeis</button>
                <div className={`${processing? 'block':'hidden'}`}>
                  <div className="progress-panel">
                    <div className="text-sm font-medium">{progressMessage}</div>
                    <div className="w-full bg-progress-track h-3 rounded mt-3 overflow-hidden"><div className="h-full bg-progress-fill" style={{width: `${progress}%`}} /></div>
                    <div className="text-xs text-muted mt-2">{progress}%</div>
                  </div>
                </div>
              </div>
            </div>
          </form>
        </div>
      </div>

      {modal.open && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white rounded-2xl p-6 w-full max-w-lg shadow-xl">
            <div className={`text-4xl ${modal.success? 'text-green-500':'text-red-500'}`}>{modal.success? '✓':'✕'}</div>
            <h3 className="text-xl font-semibold mt-2">{modal.success? 'Sucesso!':'Erro'}</h3>
            <p className="text-sm text-gray-600 mt-2">{modal.message}</p>
            <div className="mt-4 text-right"><button onClick={()=>setModal({open:false,success:false,message:''})} className="px-4 py-2 rounded bg-sky-600 text-white">Fechar</button></div>
          </div>
        </div>
      )}
      </div>
    </div>
  )
}
