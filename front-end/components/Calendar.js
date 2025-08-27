import { useState, useEffect, useRef } from 'react'

function formatDate(d){
  if(!d) return ''
  const y = d.getFullYear(); const m = String(d.getMonth()+1).padStart(2,'0'); const day = String(d.getDate()).padStart(2,'0')
  return `${y}-${m}-${day}`
}

function parseDate(s){
  if(!s) return null
  const parts = s.split('-')
  if(parts.length!==3) return null
  return new Date(Number(parts[0]), Number(parts[1])-1, Number(parts[2]))
}

export default function Calendar({ value, onChange, placeholder }){
  const [open, setOpen] = useState(false)
  const [activeMonth, setActiveMonth] = useState(()=> parseDate(value) || new Date())
  const ref = useRef(null)

  useEffect(()=>{ if(value) setActiveMonth(parseDate(value)) },[value])

  useEffect(()=>{
    function onDoc(e){ if(ref.current && !ref.current.contains(e.target)) setOpen(false) }
    document.addEventListener('click', onDoc)
    return ()=> document.removeEventListener('click', onDoc)
  },[])

  const startOfMonth = (d)=> new Date(d.getFullYear(), d.getMonth(), 1)
  const endOfMonth = (d)=> new Date(d.getFullYear(), d.getMonth()+1, 0)

  const buildGrid = (d) => {
    const start = startOfMonth(d)
    const end = endOfMonth(d)
    // weekday of first day (0=Sun..6=Sat). We'll start week on Sun like template.
    const startWeek = start.getDay()
    const days = []
    // previous month tail
    for(let i = startWeek-1; i >=0; i--){
      const prev = new Date(start.getFullYear(), start.getMonth(), -i)
      days.push({ date: prev, inMonth: false })
    }
    for(let i=1;i<=end.getDate();i++) days.push({ date: new Date(d.getFullYear(), d.getMonth(), i), inMonth: true })
    // fill to complete weeks (42 cells max)
    while(days.length % 7 !== 0) {
      const last = days[days.length-1].date
      const next = new Date(last.getFullYear(), last.getMonth(), last.getDate()+1)
      days.push({ date: next, inMonth: false })
    }
    return days
  }

  const grid = buildGrid(activeMonth)

  function selectDay(d){
    const s = formatDate(d)
    onChange && onChange(s)
    setOpen(false)
  }

  function prevMonth(){ setActiveMonth(m=> new Date(m.getFullYear(), m.getMonth()-1, 1)) }
  function nextMonth(){ setActiveMonth(m=> new Date(m.getFullYear(), m.getMonth()+1, 1)) }

  return (
    <div className="relative" ref={ref}>
      <button type="button" className="calendar-input" onClick={()=>setOpen(o=>!o)}>
        <span className="calendar-input-text">{ value ? ( (function(v){ const p = v.split('-'); return `${p[2]}/${p[1]}/${p[0]}` })(value) ) : (placeholder || 'Selecione a data') }</span>
        <span className="calendar-input-icon">▾</span>
      </button>

      {open && (
        <div className="calendar-popup">
          <div className="calendar-header">
            <button type="button" className="calendar-nav" onClick={prevMonth} aria-label="Anterior">‹</button>
            <div className="calendar-month">{activeMonth.toLocaleString('pt-BR',{ month: 'long', year: 'numeric' })}</div>
            <button type="button" className="calendar-nav" onClick={nextMonth} aria-label="Próximo">›</button>
          </div>

          <div className="calendar-grid">
            {['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'].map(d=> <div key={d} className="calendar-weekday">{d}</div>)}
            {grid.map((cell, idx) => {
              const s = formatDate(cell.date)
              const isSelected = value === s
              return (
                <button key={idx} type="button" onClick={()=>selectDay(cell.date)} className={`calendar-day ${cell.inMonth? '':'calendar-day--muted'} ${isSelected? 'calendar-day--selected':''}`}>
                  {cell.date.getDate()}
                </button>
              )
            })}
          </div>
        </div>
      )}
    </div>
  )
}
