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
} from 'recharts';
import KpiCard from '../components/KpiCard';
import FilterBar from '../components/FilterBar';
import ChartCard from '../components/ChartCard';
import {
  loadOpportunityImport,
  normalizeDashboardDate,
  OPPORTUNITY_IMPORT_EVENT,
  normalizeBusinessValue,
} from '../utils/opportunityImport';

const STATIC_CLASSEUR_URL = '/data/Classeur.xlsx';
const CLASSEUR_QUOTATION_GROUPS = [
  'Sales',
  'CRC',
  'UW',
  'Reinsurance',
  'Risk engineer',
  'Doctor',
];
const CLASSEUR_POLICY_GROUPS = ['AML', 'Finance', 'Policy admin'];

const POLICY_GWP_STATUSES = new Set([
  'aml screening pending',
  'customer creation in progress',
  'customer creation pending',
  'policy generation in-progress',
  'policy generation pending',
  'returned by finance',
  'uw clearance in-progress',
  'complete',
]);

const minFilterDate = '2026-01-01';
const maxFilterDate = '2026-06-30';

const isValidIsoDate = (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value ?? ''));

const normalizeDashboardBusiness = (value) => {
  const normalized = normalizeBusinessValue(value);

  if (normalized === 'motor') return 'Motor';
  if (normalized === 'medical') return 'Medical';
  if (normalized === 'life') return 'Life';
  if (normalized === 'general') return 'General';

  return '';
};

const normalizeStatus = (value) => String(value ?? '').trim().toLowerCase();

const normalizeGroupName = (value) =>
  String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');

const parseAmount = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  const normalized = String(value)
    .replace(/,/g, '')
    .replace(/\s/g, '')
    .trim();

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeClasseurDate = (value) => {
  const rawValue = String(value ?? '').trim();
  if (!rawValue) return '';

  if (isValidIsoDate(rawValue)) {
    return rawValue;
  }

  const match = rawValue.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})(?:\s|$)/);
  if (match) {
    return normalizeDashboardDate(`${match[1]}-${match[2]}-${match[3]}`);
  }

  return '';
};

const loadClasseurWorkbookRows = async () => {
  const response = await fetch(`${STATIC_CLASSEUR_URL}?t=${Date.now()}`, {
    cache: 'no-store',
    headers: {
      'Cache-Control': 'no-cache',
      Pragma: 'no-cache',
    },
  });

  if (!response.ok) {
    throw new Error('Classeur workbook not found');
  }

  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  return XLSX.utils.sheet_to_json(worksheet, {
    defval: '',
    raw: false,
  });
};

const buildTatChartData = (rows, labels) => {
  return labels.map((label) => {
    const normalizedLabel = normalizeGroupName(label);

    const matchingRows = rows.filter(
      (row) => normalizeGroupName(row.UsrAssignToGroup) === normalizedLabel
    );

    const count = matchingRows.length;
    const totalDays = matchingRows.reduce(
      (acc, row) => acc + parseAmount(row.UsrTurnaroundDays),
      0
    );
    const totalMinutes = matchingRows.reduce(
      (acc, row) => acc + parseAmount(row.UsrTurnaroundTime),
      0
    );

    return {
      name: label,
      count,
      avgDays: count === 0 ? 0 : totalDays / count,
      avgMinutes: count === 0 ? 0 : totalMinutes / count,
    };
  });
};

function TatTooltip({ active, payload, label }) {
  if (!active || !payload || payload.length === 0) return null;

  const data = payload[0]?.payload;
  if (!data) return null;

  return (
    <div
      style={{
        background: 'white',
        border: '1px solid #e2e8f0',
        borderRadius: 12,
        padding: 12,
        boxShadow: '0 10px 25px rgba(15, 23, 42, 0.08)',
      }}
    >
      <div style={{ fontWeight: 700, marginBottom: 8 }}>{label}</div>
      <div>Count: {data.count}</div>
      <div>Avg TAT (Days): {data.avgDays.toFixed(2)}</div>
      <div>Avg TAT (Minutes): {data.avgMinutes.toFixed(2)}</div>
    </div>
  );
}

export default function DashboardPage() {
  const [importedWorkbook, setImportedWorkbook] = useState(() => loadOpportunityImport());
  const [classeurRows, setClasseurRows] = useState([]);
  const [isClasseurLoading, setIsClasseurLoading] = useState(false);

  const importedSheet =
    importedWorkbook?.workbookData?.sheets?.find(
      (sheet) => sheet.name === importedWorkbook?.activeSheet
    ) ??
    importedWorkbook?.workbookData?.sheets?.[0] ??
    null;

  const importedRows = importedSheet?.rows ?? [];

  useEffect(() => {
    const syncImportedWorkbook = () => {
      setImportedWorkbook(loadOpportunityImport());
    };

    window.addEventListener(OPPORTUNITY_IMPORT_EVENT, syncImportedWorkbook);
    window.addEventListener('storage', syncImportedWorkbook);
    window.addEventListener('focus', syncImportedWorkbook);

    return () => {
      window.removeEventListener(OPPORTUNITY_IMPORT_EVENT, syncImportedWorkbook);
      window.removeEventListener('storage', syncImportedWorkbook);
      window.removeEventListener('focus', syncImportedWorkbook);
    };
  }, []);

  useEffect(() => {
    const loadClasseur = async () => {
      try {
        setIsClasseurLoading(true);
        const rows = await loadClasseurWorkbookRows();
        setClasseurRows(rows);
      } catch (error) {
        console.error('Failed to load Classeur.xlsx', error);
        setClasseurRows([]);
      } finally {
        setIsClasseurLoading(false);
      }
    };

    loadClasseur();
  }, []);

  const allOppDates = importedRows
    .map((row) => normalizeDashboardDate(row.UsrQuoteSubmissionDate))
    .filter((value) => isValidIsoDate(value));

  const classeurDates = classeurRows
    .map((row) => normalizeClasseurDate(row.UsrRequestDateTimeIn))
    .filter((value) => isValidIsoDate(value));

  const combinedDates = [...allOppDates, ...classeurDates].sort();

  const effectiveMinDate = combinedDates[0] ?? minFilterDate;
  const effectiveMaxDate = combinedDates[combinedDates.length - 1] ?? maxFilterDate;

  const [filters, setFilters] = useState({
    fromDate: effectiveMinDate,
    toDate: effectiveMaxDate,
    businesses: [],
  });

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
        return {
          ...prev,
          fromDate: effectiveMinDate,
          toDate: effectiveMaxDate,
        };
      }

      if (nextFrom === prev.fromDate && nextTo === prev.toDate) {
        return prev;
      }

      return {
        ...prev,
        fromDate: nextFrom,
        toDate: nextTo,
      };
    });
  }, [effectiveMinDate, effectiveMaxDate]);

  const handleFilterChange = (key, value) => {
    setFilters((prev) => ({
      ...prev,
      [key]: value,
    }));
  };

  const filteredImportedRows = useMemo(() => {
    if (importedRows.length === 0) return [];

    return importedRows.filter((row) => {
      const submissionDate = normalizeDashboardDate(row.UsrQuoteSubmissionDate);
      if (!isValidIsoDate(submissionDate)) {
        return false;
      }

      const matchesDate =
        submissionDate >= filters.fromDate && submissionDate <= filters.toDate;

      const mappedBusiness = normalizeDashboardBusiness(
        row['Business Mapping'] ?? row.UsrClass
      );

      const matchesBusiness =
        !Array.isArray(filters.businesses) ||
        filters.businesses.length === 0 ||
        filters.businesses.includes(mappedBusiness);

      return matchesDate && matchesBusiness;
    });
  }, [filters.businesses, filters.fromDate, filters.toDate, importedRows]);

  const filteredClasseurRows = useMemo(() => {
    if (classeurRows.length === 0) return [];

    return classeurRows.filter((row) => {
      const requestDate = normalizeClasseurDate(row.UsrRequestDateTimeIn);
      if (!isValidIsoDate(requestDate)) {
        return false;
      }

      const matchesDate =
        requestDate >= filters.fromDate && requestDate <= filters.toDate;

      const mappedBusiness = normalizeDashboardBusiness(row.UsrBusiness);

      const matchesBusiness =
        !Array.isArray(filters.businesses) ||
        filters.businesses.length === 0 ||
        filters.businesses.includes(mappedBusiness);

      return matchesDate && matchesBusiness;
    });
  }, [classeurRows, filters.businesses, filters.fromDate, filters.toDate]);

  const quotationsCount = filteredImportedRows.length;

  const policiesConvertedCount = useMemo(() => {
    return filteredImportedRows.filter(
      (row) => normalizeStatus(row.UsrStatus) === 'complete'
    ).length;
  }, [filteredImportedRows]);

  const quotationsGwp = useMemo(() => {
    return filteredImportedRows.reduce(
      (acc, row) => acc + parseAmount(row.UsrQuotationAmount),
      0
    );
  }, [filteredImportedRows]);

  const policiesGwp = useMemo(() => {
    return filteredImportedRows
      .filter((row) => POLICY_GWP_STATUSES.has(normalizeStatus(row.UsrStatus)))
      .reduce((acc, row) => acc + parseAmount(row.UsrQuotationAmount), 0);
  }, [filteredImportedRows]);

  const metrics = useMemo(() => {
    const quotations = quotationsCount;
    const policies = policiesConvertedCount;
    const conversionRate = quotations === 0 ? 0 : (policies / quotations) * 100;
    const gwpConversionRate =
      quotationsGwp === 0 ? 0 : (policiesGwp / quotationsGwp) * 100;

    return {
      quotations,
      policies,
      conversionRate,
      quotationsGwp,
      policiesGwp,
      gwpConversionRate,
    };
  }, [quotationsCount, policiesConvertedCount, quotationsGwp, policiesGwp]);

  const quotationsTatByDepartment = useMemo(() => {
    return buildTatChartData(filteredClasseurRows, CLASSEUR_QUOTATION_GROUPS);
  }, [filteredClasseurRows]);

  const policiesTatByDepartment = useMemo(() => {
    return buildTatChartData(filteredClasseurRows, CLASSEUR_POLICY_GROUPS);
  }, [filteredClasseurRows]);

  return (
    <>
      <FilterBar filters={filters} onFilterChange={handleFilterChange} />

      <section className="kpi-grid kpi-grid-three">
        <KpiCard title="Quotations" value={metrics.quotations} showTarget={false} />
        <KpiCard title="Policies Converted" value={metrics.policies} showTarget={false} />
        <KpiCard
          title="Conversion Rate"
          value={metrics.conversionRate}
          suffix="%"
          showTarget={false}
        />
      </section>

      <section className="kpi-grid kpi-grid-three">
        <KpiCard
          title="Quotations GWP"
          value={metrics.quotationsGwp}
          currency
          showTarget={false}
        />
        <KpiCard
          title="Policies GWP"
          value={metrics.policiesGwp}
          currency
          showTarget={false}
        />
        <KpiCard
          title="GWP Conversion Rate"
          value={metrics.gwpConversionRate}
          suffix="%"
          showTarget={false}
        />
      </section>

      <section className="two-col">
        <ChartCard title="Quotations TAT by Department">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={quotationsTatByDepartment}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" interval={0} angle={-10} textAnchor="end" height={60} />
              <YAxis />
              <Tooltip content={<TatTooltip />} />
              <Bar dataKey="count" name="Count" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Policies TAT by Department">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={policiesTatByDepartment}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" interval={0} angle={-10} textAnchor="end" height={60} />
              <YAxis />
              <Tooltip content={<TatTooltip />} />
              <Bar dataKey="count" name="Count" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>

      {isClasseurLoading ? (
        <section className="two-col">
          <ChartCard title="Loading Classeur TAT Data">
            <div style={{ padding: '24px 8px' }}>Loading Classeur.xlsx...</div>
          </ChartCard>
        </section>
      ) : null}
    </>
  );
}
