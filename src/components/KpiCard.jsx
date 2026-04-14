export default function KpiCard({ title, value, target, suffix = '', currency = false, showTarget = true }) {
  const hasTarget = typeof target === 'number';
  const variance = hasTarget ? value - target : 0;
  const variancePct = !hasTarget || target === 0 ? 0 : (variance / target) * 100;

  const formatValue = (num) => {
    if (currency) {
      return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(num);
    }
    return new Intl.NumberFormat('en-US', { maximumFractionDigits: 1 }).format(num);
  };

  return (
    <div className="card kpi-card">
      <div className="kpi-label">{title}</div>
      <div className="kpi-value">{currency ? 'SAR ' : ''}{formatValue(value)}{suffix}</div>
      {hasTarget ? (
        <>
          {showTarget ? <div className="kpi-meta">Target: {currency ? 'SAR ' : ''}{formatValue(target)}{suffix}</div> : null}
          <div className={`variance ${variance >= 0 ? 'positive' : 'negative'}`}>
            {variance >= 0 ? '+' : ''}{variancePct.toFixed(1)}%
          </div>
        </>
      ) : null}
    </div>
  );
}
