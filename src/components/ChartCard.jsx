export default function ChartCard({ title, children }) {
  return (
    <div className="card chart-card">
      <div className="section-title">{title}</div>
      <div className="chart-wrap">{children}</div>
    </div>
  );
}
