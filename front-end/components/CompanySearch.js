import { useState, useEffect } from 'react'
import { Combobox } from '@headlessui/react'

export default function CompanySearch({ value, onChange }){
  const [query, setQuery] = useState('')
  const [companies, setCompanies] = useState([])

  useEffect(() => {
    // mock data; replace with API
    const sample = [
      { id: 1, name: 'Empresa Alpha' },
      { id: 2, name: 'Beta Comércio' },
      { id: 3, name: 'Gamma Ltda' },
      { id: 4, name: 'Delta Serviços' }
    ]
    setCompanies(sample)
  }, [])

  const filtered = query === '' ? companies : companies.filter(c => c.name.toLowerCase().includes(query.toLowerCase()))

  return (
    <Combobox value={value} onChange={onChange}>
      <div className="relative">
        <Combobox.Input className="w-full px-3 py-2 border rounded" displayValue={c => c ? c.name : ''} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar empresa..." />
        {filtered.length > 0 && (
          <Combobox.Options className="absolute z-20 mt-1 w-full bg-white border rounded shadow-sm max-h-48 overflow-auto">
            {filtered.map(c => (
              <Combobox.Option key={c.id} value={c} className={({active}) => `px-3 py-2 cursor-pointer ${active ? 'bg-slate-50' : ''}`}>
                {c.name}
              </Combobox.Option>
            ))}
          </Combobox.Options>
        )}
      </div>
    </Combobox>
  )
}
