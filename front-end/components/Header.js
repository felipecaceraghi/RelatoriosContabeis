export default function Header(){
  return (
    <header className="header mb-6">
      <div className="container">
        <div className="flex items-center gap-4">
          <img src="/logo.png" alt="logo" className="header-logo" />
          <div>
            <div className="text-lg font-semibold">Sistema de Automação - Relatórios Contábeis</div>
            <div className="text-sm text-white/80">Painel administrativo</div>
          </div>
        </div>
      </div>
    </header>
  )
}
