export default function FormCard({ children, title, subtitle }){
  return (
    <section className="card card-panel">
      <div className="card-header">
        <h2 className="card-title">{title}</h2>
        {subtitle && <p className="card-subtitle">{subtitle}</p>}
      </div>
      <div className="card-body">{children}</div>
    </section>
  )
}
