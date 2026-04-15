import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  Legend,
  PieChart,
  Pie,
  Cell,
} from 'recharts';
import KpiCard from '../components/KpiCard';
import FilterBar from '../components/FilterBar';
import ChartCard from '../components/ChartCard';
import { normalizeBusinessValue, normalizeDashboardDate } from '../utils/opportunityImport';

const STATIC_ALL_OPP_URL = '/data/AllOppQDR.xlsx';
const STATIC_CLASSEUR_URL = '/data/Classeur.xlsx';

const minFilterDate = '2026-01-01';
const maxFilterDate = '2026-06-30';

const POLICY_GWP_STATUSES = new Set([
  'aml screening pending',
  'customer creation in-progress',
  'customer creation in progress',
  'customer creation pending',
  'escalated (at the account opening stage)',
  'policy generation in-progress',
  'policy generation pending',
  'returned by finance',
  'uw clearance in-progress',
  'complete',
]);

const QUOTATION_TAT_GROUPS = ['Sales', 'UW', 'Reinsurance', 'Risk engineer', 'Doctor'];
const POLICY_TAT_GROUPS = ['AML', 'Finance', 'PA'];
const ACTION_COLORS = ['#2563eb', '#16a34a', '#dc2626'];

const isValidIsoDate = (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value ?? ''));

const normalizeDashboardBusiness = (value) => {
  const normalized = normalizeBusinessValue(value);
  if (normalized === 'motor') return 'Motor';
  if (normalized === 'medical') return 'Medical';
  if (normalized === 'life') return 'Life';
  if (normalized === 'general') return 'General';
  return '';
};

const normalizeBusinessFromUsrClass = (usrClassValue) => {
  const normalized = normalizeBusinessValue(usrClassValue);
  if (normalized === 'motor') return 'Motor';
  if (normalized === 'medical') return 'Medical';
  if (normalized === 'life') return 'Life';
  if (normalized === 'general') return 'General';
  return '';
};

const parseAmount = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const normalized = String(value).replace(/,/g, '').replace(/\s/g, '').trim();
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
};

const parseTurnaroundDays = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const raw = String(value).trim();
  const parsed = Number(raw.replace(/,/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
};

const parseTurnaroundMinutes = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  const raw = String(value).trim();

  const mmssMatch = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (mmssMatch) {
    const minutes = Number(mmssMatch[1]);
    const seconds = Number(mmssMatch[2]);
    if (Number.isFinite(minutes) && Number.isFinite(seconds)) {
      return minutes + seconds / 60;
    }
  }

  const hhmmssMatch = raw.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (hhmmssMatch) {
    const hours = Number(hhmmssMatch[1]);
    const minutes = Number(hhmmssMatch[2]);
    const seconds = Number(hhmmssMatch[3]);
    if (Number.isFinite(hours) && Number.isFinite(minutes) && Number.isFinite(seconds)) {
      return hours * 60 + minutes + seconds / 60;
    }
  }

  const parsed = Number(raw.replace(/,/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeStatus = (value) => String(value ?? '').trim().toLowerCase();

const normalizeAssignedByRoleForQuotationTat = (value) => {
  const raw = String(value ?? '').trim().toLowerCase();
  if (!raw) return '';
  if (raw.includes('corporate sales')) return 'Sales';
  if (raw.includes('underwriting')) return 'UW';
  if (raw.includes('reinsurance') || raw.includes('reinusrance')) return 'Reinsurance';
  if (raw.includes('risk engineer') || raw.includes('riskengineer')) return 'Risk engineer';
  if (raw.includes('doctor')) return 'Doctor';
  return '';
};

const normalizeAssignedByRoleForPoliciesTat = (value) => {
  const raw = String(value ?? '').trim().toLowerCase();
  if (!raw) return '';
  if (raw.includes('aml')) return 'AML';
  if (raw.includes('finance')) return 'Finance';
  if (
    raw === 'pa' ||
    raw.includes('policy admin') ||
    raw.includes('policy administration') ||
    raw.includes('policy admin member') ||
    raw.includes('policy admin agent') ||
    raw.includes('policy admin supervisor') ||
    raw.includes('policy issuance') ||
    raw.includes('policy')
  ) {
    return 'PA';
  }
  return '';
};

const normalizeAction = (value) => {
  const raw = String(value ?? '').trim().toLowerCase();
  if (!raw) return '';
  if (raw === 'approve' || raw.includes('approve')) {
    if (raw.includes('conditional')) return 'Approve Conditionally';
    return 'Approve';
  }
  if (raw.includes('reject')) return 'Reject';
  return '';
};

const normalizeClasseurDate = (value) => {
  const rawValue = String(value ?? '').trim();
  if (!rawValue) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(rawValue)) return rawValue;

  const match = rawValue.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})/);
  if (!match) return '';

  const [, day, monthToken, yearToken] = match;
  const monthMap = {
    JAN: '01',
    FEB: '02',
    MAR: '03',
    APR: '04',
    MAY: '05',
    JUN: '06',
    JUL: '07',
    AUG: '08',
    SEP: '09',
    OCT: '10',
    NOV: '11',
    DEC: '12',
  };

  const month = monthMap[monthToken.toUpperCase()];
  if (!month) return '';

  const year = yearToken.length === 2 ? `20${yearToken}` : yearToken.padStart(4, '0');
  return `${year}-${month}-${day.padStart(2, '0')}`;
};

const formatMillions = (value) => {
  const num = Number(value) || 0;
  if (Math.abs(num) >= 1_000_000_000) return `${(num / 1_000_000_000).toFixed(1)}B`;
  if (Math.abs(num) >= 1_000_000) return `${(num / 1_000_000).toFixed(0)}M`;
  if (Math.abs(num) >= 1_000) return `${(num / 1_000).toFixed(0)}K`;
  return `${num}`;
};

const formatSourceAxisLabel = (value) => {
  const raw = String(value ?? '').trim();
  if (raw.toLowerCase() === 'corporate sales. director/ country manager') return 'Corporate Sales';
  if (raw.length > 18) return `${raw.slice(0, 16)}…`;
  return raw;
};

const buildQuotationTatDataset = (rows, groups) =>
  groups.map((groupName) => {
    const groupRows = rows.filter(
      (row) => normalizeAssignedByRoleForQuotationTat(row.UsrAssignedByRole) === groupName
    );

    const rowsWithRequestId = groupRows.filter(
      (row) => String(row.UsrRequestNoId ?? '').trim() !== ''
    );

    const count = rowsWithRequestId.length;

    const avgDays =
      count === 0
        ? 0
        : rowsWithRequestId.reduce(
            (acc, row) => acc + parseTurnaroundDays(row.UsrTurnaroundDays),
            0
          ) / count;

    const avgMinutes =
      count === 0
        ? 0
        : rowsWithRequestId.reduce(
            (acc, row) => acc + parseTurnaroundMinutes(row.UsrTurnaroundTime),
            0
          ) / count;

    return {
      name: groupName,
      count,
      avgDays: Number(avgDays.toFixed(2)),
      avgMinutes: Number(avgMinutes.toFixed(2)),
    };
  });

const buildPolicyTatDataset = (rows, groups) =>
  groups.map((groupName) => {
    const groupRows = rows.filter(
      (row) => normalizeAssignedByRoleForPoliciesTat(row.UsrAssignedByRole) === groupName
    );

    const count = groupRows.length;

    const avgDays =
      count === 0
        ? 0
        : groupRows.reduce((acc, row) => acc + parseTurnaroundDays(row.UsrTurnaroundDays), 0) /
          count;

    const avgMinutes =
      count === 0
        ? 0
        : groupRows.reduce(
            (acc, row) => acc + parseTurnaroundMinutes(row.UsrTurnaroundTime),
            0
          ) / count;

    return {
      name: groupName,
      count,
      avgDays: Number(avgDays.toFixed(2)),
      avgMinutes: Number(avgMinutes.toFixed(2)),
    };
  });

const TatTooltip = ({ active, payload, label }) => {
  if (!active || !payload || payload.length === 0) return null;
  const data = payload[0]?.payload;
  if (!data) return null;

  return (
    <div style={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: 12, padding: 12, boxShadow: '0 8px 24px rgba(15,23,42,0.08)' }}>
      <div style={{ fontWeight: 700, marginBottom: 6 }}>{label}</div>
      <div>Count: {data.count}</div>
      <div>Avg TAT Days: {data.avgDays}</div>
      <div>Avg TAT Minutes: {data.avgMinutes}</div>
    </div>
  );
};

const ConversionRateTooltip = ({ active, payload, label }) => {
  if (!active || !payload || payload.length === 0) return null;
  const data = payload[0]?.payload;
  if (!data) return null;

  return (
    <div style={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: 12, padding: 12, boxShadow: '0 8px 24px rgba(15,23,42,0.08)' }}>
      <div style={{ fontWeight: 700, marginBottom: 6 }}>{label}</div>
      <div>Quotations: {data.quotations}</div>
      <div>Complete: {data.policies}</div>
      <div>Conversion Rate: {data.value}%</div>
    </div>
  );
};

const ActionPieTooltip = ({ active, payload }) => {
  if (!active || !payload || payload.length === 0) return null;
  const data = payload[0]?.payload;
  if (!data) return null;

  return (
    <div style={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: 12, padding: 12, boxShadow: '0 8px 24px rgba(15,23,42,0.08)' }}>
      <div style={{ fontWeight: 700, marginBottom: 6 }}>{data.name}</div>
      <div>Count: {data.value}</div>
    </div>
  );
};

const GwpTooltip = ({ active, payload, label }) => {
  if (!active || !payload || payload.length === 0) return null;

  return (
    <div style={{ background: 'white', border: '1px solid #e2e8f0', borderRadius: 12, padding: 12, boxShadow: '0 8px 24px rgba(15,23,42,0.08)' }}>
      <div style={{ fontWeight: 700, marginBottom: 6 }}>{label}</div>
      {payload.map((entry) => (
        <div key={entry.dataKey}>
          {entry.name}: SAR {Number(entry.value).toLocaleString()}
        </div>
      ))}
    </div>
  );
};

export default function DashboardPage() {
  const [allOppRows, setAllOppRows] = useState([]);
  const [classeurRows, setClasseurRows] = useState([]);
  const [allOppError, setAllOppError] = useState('');
  const [classeurError, setClasseurError] = useState('');

  const allOppDates = allOppRows
    .map((row) => normalizeDashboardDate(row.UsrQuoteSubmissionDate))
    .filter((value) => isValidIsoDate(value))
    .sort();

  const effectiveMinDate = allOppDates[0] ?? minFilterDate;
  const effectiveMaxDate = allOppDates[allOppDates.length - 1] ?? maxFilterDate;

  const [filters, setFilters] = useState({
    fromDate: minFilterDate,
    toDate: maxFilterDate,
    businesses: [],
  });

  useEffect(() => {
    const loadAllOppWorkbook = async () => {
      try {
        setAllOppError('');
        const response = await fetch(`${STATIC_ALL_OPP_URL}?t=${Date.now()}`, {
          cache: 'no-store',
          headers: { 'Cache-Control': 'no-cache', Pragma: 'no-cache' },
        });

        if (!response.ok) throw new Error('AllOppQDR workbook not found');

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '', raw: false });

        const normalizedRows = rawRows.map((row) => ({
          ...row,
          UsrQuoteSubmissionDate: normalizeDashboardDate(row.UsrQuoteSubmissionDate),
          'Business Mapping':
            String(row['Business Mapping'] ?? '').trim() ||
            normalizeBusinessFromUsrClass(row.UsrClass),
        }));

        setAllOppRows(normalizedRows);
      } catch {
        setAllOppRows([]);
        setAllOppError('AllOppQDR workbook could not be loaded.');
      }
    };

    loadAllOppWorkbook();
  }, []);

  useEffect(() => {
    const loadClasseurWorkbook = async () => {
      try {
        setClasseurError('');
        const response = await fetch(`${STATIC_CLASSEUR_URL}?t=${Date.now()}`, {
          cache: 'no-store',
          headers: { 'Cache-Control': 'no-cache', Pragma: 'no-cache' },
        });

        if (!response.ok) throw new Error('Classeur workbook not found');

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '', raw: false });

        setClasseurRows(rawRows);
      } catch {
        setClasseurRows([]);
        setClasseurError('Classeur workbook could not be loaded.');
      }
    };

    loadClasseurWorkbook();
  }, []);

  useEffect(() => {
    setFilters((prev) => {
      const nextFrom =
        !prev.fromDate || prev.fromDate < effectiveMinDate || prev.fromDate > effectiveMaxDate
          ? effectiveMinDate
          : prev.fromDate;

      const nextTo =
        !prev.toDate || prev.toDate < effectiveMinDate || prev.toDate > effectiveMaxDate
          ? effectiveMaxDate
          : prev.toDate;

      if (nextFrom > nextTo) {
        return { ...prev, fromDate: effectiveMinDate, toDate: effectiveMaxDate };
      }

      if (nextFrom === prev.fromDate && nextTo === prev.toDate) return prev;
      return { ...prev, fromDate: nextFrom, toDate: nextTo };
    });
  }, [effectiveMinDate, effectiveMaxDate]);

  const handleFilterChange = (key, value) => {
    setFilters((prev) => ({ ...prev, [key]: value }));
  };

  const filteredImportedRows = useMemo(() => {
    if (allOppRows.length === 0) return [];

    return allOppRows.filter((row) => {
      const submissionDate = normalizeDashboardDate(row.UsrQuoteSubmissionDate);
      if (!isValidIsoDate(submissionDate)) return false;

      const matchesDate = submissionDate >= filters.fromDate && submissionDate <= filters.toDate;
      const mappedBusiness = normalizeDashboardBusiness(
        row['Business Mapping'] ?? row.UsrClass
      );

      const matchesBusiness =
        !Array.isArray(filters.businesses) ||
        filters.businesses.length === 0 ||
        filters.businesses.includes(mappedBusiness);

      return matchesDate && matchesBusiness;
    });
  }, [allOppRows, filters.businesses, filters.fromDate, filters.toDate]);

  const filteredClasseurRows = useMemo(() => {
    if (classeurRows.length === 0) return [];

    return classeurRows.filter((row) => {
      const requestDate = normalizeClasseurDate(row.UsrRequestDateTimeIn);
      if (!isValidIsoDate(requestDate)) return false;

      const matchesDate = requestDate >= filters.fromDate && requestDate <= filters.toDate;
      const mappedBusiness = normalizeDashboardBusiness(row.UsrBusiness);

      const matchesBusiness =
        !Array.isArray(filters.businesses) ||
        filters.businesses.length === 0 ||
        filters.businesses.includes(mappedBusiness);

      return matchesDate && matchesBusiness;
    });
  }, [classeurRows, filters.businesses, filters.fromDate, filters.toDate]);

  const quotationsTatByDepartment = useMemo(
    () => buildQuotationTatDataset(filteredClasseurRows, QUOTATION_TAT_GROUPS),
    [filteredClasseurRows]
  );

  const policiesTatByDepartment = useMemo(
    () => buildPolicyTatDataset(filteredClasseurRows, POLICY_TAT_GROUPS),
    [filteredClasseurRows]
  );

  const actionPieData = useMemo(() => {
    const grouped = filteredClasseurRows.reduce((acc, row) => {
      const action = normalizeAction(row.UsrAction);
      if (!action) return acc;
      if (!acc[action]) acc[action] = 0;
      acc[action] += 1;
      return acc;
    }, {});

    const ordered = ['Approve', 'Approve Conditionally', 'Reject'];

    return ordered
      .filter((name) => grouped[name] > 0)
      .map((name) => ({ name, value: grouped[name] }));
  }, [filteredClasseurRows]);

  const quotationsCount = useMemo(() => filteredImportedRows.length, [filteredImportedRows]);

  const policiesConvertedCount = useMemo(
    () =>
      filteredImportedRows.filter((row) => normalizeStatus(row.UsrStatus) === 'complete').length,
    [filteredImportedRows]
  );

  const quotationsGwp = useMemo(
    () =>
      filteredImportedRows.reduce((acc, row) => acc + parseAmount(row.UsrQuotationAmount), 0),
    [filteredImportedRows]
  );

  const policiesGwp = useMemo(
    () =>
      filteredImportedRows
        .filter((row) => POLICY_GWP_STATUSES.has(normalizeStatus(row.UsrStatus)))
        .reduce((acc, row) => acc + parseAmount(row.UsrQuotationAmount), 0),
    [filteredImportedRows]
  );

  const metrics = useMemo(() => {
    const quotations = quotationsCount;
    const policies = policiesConvertedCount;
    const conversionRate = quotations === 0 ? 0 : (policies / quotations) * 100;
    const gwpConversionRate = quotationsGwp === 0 ? 0 : (policiesGwp / quotationsGwp) * 100;

    return {
      quotations,
      policies,
      conversionRate,
      quotationsGwp,
      policiesGwp,
      gwpConversionRate,
    };
  }, [quotationsCount, policiesConvertedCount, quotationsGwp, policiesGwp]);

  const quotesByStatus = useMemo(() => {
    const grouped = filteredImportedRows.reduce((acc, row) => {
      const status = String(row.UsrStatus ?? '').trim() || 'Blank';
      acc[status] = (acc[status] || 0) + 1;
      return acc;
    }, {});

    return Object.entries(grouped)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
  }, [filteredImportedRows]);

  const conversionByBusiness = useMemo(() => {
    const grouped = filteredImportedRows.reduce((acc, row) => {
      const business =
        normalizeDashboardBusiness(row['Business Mapping'] ?? row.UsrClass) || 'Unknown';

      if (!acc[business]) acc[business] = { quotations: 0, policies: 0 };
      acc[business].quotations += 1;

      if (normalizeStatus(row.UsrStatus) === 'complete') {
        acc[business].policies += 1;
      }

      return acc;
    }, {});

    return Object.entries(grouped).map(([name, val]) => ({
      name,
      quotations: val.quotations,
      policies: val.policies,
      value:
        val.quotations === 0
          ? 0
          : Number(((val.policies / val.quotations) * 100).toFixed(1)),
    }));
  }, [filteredImportedRows]);

  const gwpByRegionQuotationVsPolicies = useMemo(() => {
    const grouped = filteredImportedRows.reduce((acc, row) => {
      const region = String(row.UsrRegion ?? '').trim() || 'Blank';
      if (!acc[region]) acc[region] = { name: region, quotation: 0, policies: 0 };

      const amount = parseAmount(row.UsrQuotationAmount);
      acc[region].quotation += amount;
      if (POLICY_GWP_STATUSES.has(normalizeStatus(row.UsrStatus))) {
        acc[region].policies += amount;
      }

      return acc;
    }, {});

    return Object.values(grouped).sort((a, b) => b.quotation - a.quotation);
  }, [filteredImportedRows]);

  const gwpByBusinessQuotationVsPolicies = useMemo(() => {
    const grouped = filteredImportedRows.reduce((acc, row) => {
      const business =
        normalizeDashboardBusiness(row['Business Mapping'] ?? row.UsrClass) || 'Unknown';

      if (!acc[business]) acc[business] = { name: business, quotation: 0, policies: 0 };

      const amount = parseAmount(row.UsrQuotationAmount);
      acc[business].quotation += amount;
      if (POLICY_GWP_STATUSES.has(normalizeStatus(row.UsrStatus))) {
        acc[business].policies += amount;
      }

      return acc;
    }, {});

    return Object.values(grouped).sort((a, b) => b.quotation - a.quotation);
  }, [filteredImportedRows]);

  const gwpBySourceQuotationVsPolicies = useMemo(() => {
    const grouped = filteredImportedRows.reduce((acc, row) => {
      const source = String(row.UsrSource ?? '').trim() || 'Blank';
      if (!acc[source]) acc[source] = { name: source, quotation: 0, policies: 0 };

      const amount = parseAmount(row.UsrQuotationAmount);
      acc[source].quotation += amount;
      if (POLICY_GWP_STATUSES.has(normalizeStatus(row.UsrStatus))) {
        acc[source].policies += amount;
      }

      return acc;
    }, {});

    return Object.values(grouped).sort((a, b) => b.quotation - a.quotation);
  }, [filteredImportedRows]);

  const policiesConvertedByDuration = useMemo(() => {
    const today = new Date();
    const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    const buckets = {
      '<1 day': 0,
      '1-3 days': 0,
      '4-7 days': 0,
      '>7 days': 0,
    };

    filteredClasseurRows
      .filter((row) => String(row.UsrAction ?? '').trim().toLowerCase() === 'issue policy')
      .forEach((row) => {
        const isoDate = normalizeClasseurDate(row.UsrRequestDateTimeIn);
        if (!isValidIsoDate(isoDate)) return;

        const requestDate = new Date(`${isoDate}T00:00:00`);
        const diffMs = startOfToday.getTime() - requestDate.getTime();
        const diffDays = diffMs / (1000 * 60 * 60 * 24);

        if (diffDays < 1) {
          buckets['<1 day'] += 1;
        } else if (diffDays <= 3) {
          buckets['1-3 days'] += 1;
        } else if (diffDays <= 7) {
          buckets['4-7 days'] += 1;
        } else {
          buckets['>7 days'] += 1;
        }
      });

    return [
      { name: '<1 day', value: buckets['<1 day'] },
      { name: '1-3 days', value: buckets['1-3 days'] },
      { name: '4-7 days', value: buckets['4-7 days'] },
      { name: '>7 days', value: buckets['>7 days'] },
    ];
  }, [filteredClasseurRows]);

  return (
    <>
      <FilterBar filters={filters} onFilterChange={handleFilterChange} />

      {(allOppError || classeurError) ? (
        <section className="one-col">
          <div className="card" style={{ padding: 20 }}>
            {allOppError ? <div><strong>{allOppError}</strong></div> : null}
            {classeurError ? <div><strong>{classeurError}</strong></div> : null}
          </div>
        </section>
      ) : null}

      <section className="kpi-grid kpi-grid-three">
        <KpiCard title="Quotations" value={metrics.quotations} showTarget={false} />
        <KpiCard title="Policies Converted" value={metrics.policies} showTarget={false} />
        <KpiCard title="Conversion Rate" value={metrics.conversionRate} suffix="%" showTarget={false} />
      </section>

      <section className="kpi-grid kpi-grid-three">
        <KpiCard title="Quotations GWP" value={metrics.quotationsGwp} currency showTarget={false} />
        <KpiCard title="Policies GWP" value={metrics.policiesGwp} currency showTarget={false} />
        <KpiCard title="GWP Conversion Rate" value={metrics.gwpConversionRate} suffix="%" showTarget={false} />
      </section>

      <section className="two-col">
        <ChartCard title="Quotations by Status">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={quotesByStatus} barCategoryGap="20%">
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis
                dataKey="name"
                interval={0}
                angle={-35}
                textAnchor="end"
                height={95}
                tick={{ fontSize: 12 }}
                axisLine={false}
                tickLine={false}
              />
              <YAxis tick={{ fontSize: 12 }} axisLine={false} tickLine={false} />
              <Tooltip
                formatter={(value) => [value, 'Count']}
                contentStyle={{ borderRadius: '10px', border: '1px solid #e2e8f0' }}
              />
              <Bar dataKey="value" radius={[10, 10, 0, 0]} barSize={34} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Conversion Rate by Business">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={conversionByBusiness} barCategoryGap="25%">
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis dataKey="name" tick={{ fontSize: 12 }} axisLine={false} tickLine={false} />
              <YAxis unit="%" tick={{ fontSize: 12 }} axisLine={false} tickLine={false} />
              <Tooltip content={<ConversionRateTooltip />} />
              <Bar dataKey="value" radius={[10, 10, 0, 0]} barSize={40} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>

      <section className="two-col">
        <ChartCard title="Quotations TAT by Department">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={quotationsTatByDepartment}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" interval={0} angle={-10} textAnchor="end" height={60} />
              <YAxis />
              <Tooltip content={<TatTooltip />} />
              <Bar dataKey="count" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Policies TAT by Department">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={policiesTatByDepartment}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" interval={0} angle={-10} textAnchor="end" height={60} />
              <YAxis />
              <Tooltip content={<TatTooltip />} />
              <Bar dataKey="count" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>

      {actionPieData.length > 0 ? (
        <section className="one-col">
          <ChartCard title="Classeur Actions">
            <ResponsiveContainer width="100%" height={340}>
              <PieChart>
                <Pie
                  data={actionPieData}
                  dataKey="value"
                  nameKey="name"
                  cx="50%"
                  cy="50%"
                  outerRadius={120}
                  innerRadius={60}
                  paddingAngle={3}
                  label={({ name, percent }) =>
                    percent > 0 ? `${name} ${(percent * 100).toFixed(0)}%` : ''
                  }
                >
                  {actionPieData.map((entry, index) => (
                    <Cell key={entry.name} fill={ACTION_COLORS[index % ACTION_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip content={<ActionPieTooltip />} />
                <Legend verticalAlign="bottom" height={36} />
              </PieChart>
            </ResponsiveContainer>
          </ChartCard>
        </section>
      ) : null}

      <section className="two-col">
        <ChartCard title="Policies Converted by Duration">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={policiesConvertedByDuration} barCategoryGap="28%">
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis
                dataKey="name"
                tick={{ fontSize: 13, fontWeight: 600 }}
                axisLine={false}
                tickLine={false}
              />
              <YAxis
                allowDecimals={false}
                tick={{ fontSize: 12 }}
                axisLine={false}
                tickLine={false}
              />
              <Tooltip
                formatter={(value) => [value, 'Count']}
                contentStyle={{ borderRadius: '10px', border: '1px solid #e2e8f0' }}
              />
              <Bar dataKey="value" radius={[10, 10, 0, 0]} barSize={52} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="GWP by Region: Quotation vs Policies">
          <ResponsiveContainer width="100%" height={340}>
            <BarChart data={gwpByRegionQuotationVsPolicies} barGap={10} barCategoryGap="22%">
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis
                dataKey="name"
                tick={{ fontSize: 12, fontWeight: 600 }}
                axisLine={false}
                tickLine={false}
              />
              <YAxis
                tickFormatter={formatMillions}
                tick={{ fontSize: 12 }}
                axisLine={false}
                tickLine={false}
              />
              <Legend wrapperStyle={{ paddingTop: 8 }} />
              <Tooltip content={<GwpTooltip />} />
              <Bar dataKey="quotation" name="Quotation GWP" radius={[10, 10, 0, 0]} barSize={30} />
              <Bar dataKey="policies" name="Policies GWP" radius={[10, 10, 0, 0]} barSize={30} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>

      <section className="two-col">
        <ChartCard title="GWP by Business: Quotation vs Policies">
          <ResponsiveContainer width="100%" height={340}>
            <BarChart data={gwpByBusinessQuotationVsPolicies} barGap={10} barCategoryGap="22%">
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis
                dataKey="name"
                tick={{ fontSize: 12, fontWeight: 600 }}
                axisLine={false}
                tickLine={false}
              />
              <YAxis
                tickFormatter={formatMillions}
                tick={{ fontSize: 12 }}
                axisLine={false}
                tickLine={false}
              />
              <Legend wrapperStyle={{ paddingTop: 8 }} />
              <Tooltip content={<GwpTooltip />} />
              <Bar dataKey="quotation" name="Quotation GWP" radius={[10, 10, 0, 0]} barSize={30} />
              <Bar dataKey="policies" name="Policies GWP" radius={[10, 10, 0, 0]} barSize={30} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="GWP by Source: Quotations vs Policies">
          <ResponsiveContainer width="100%" height={380}>
            <BarChart data={gwpBySourceQuotationVsPolicies} barGap={10} barCategoryGap="22%">
              <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
              <XAxis
                dataKey="name"
                interval={0}
                angle={-25}
                textAnchor="end"
                height={80}
                tickFormatter={formatSourceAxisLabel}
                tick={{ fontSize: 12, fontWeight: 600 }}
                axisLine={false}
                tickLine={false}
              />
              <YAxis
                tickFormatter={formatMillions}
                tick={{ fontSize: 12 }}
                axisLine={false}
                tickLine={false}
              />
              <Legend wrapperStyle={{ paddingTop: 8 }} />
              <Tooltip content={<GwpTooltip />} />
              <Bar dataKey="quotation" name="Quotation GWP" radius={[10, 10, 0, 0]} barSize={30} />
              <Bar dataKey="policies" name="Policies GWP" radius={[10, 10, 0, 0]} barSize={30} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>
    </>
  );
}
