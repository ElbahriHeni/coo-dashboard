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
import { records } from '../data/dashboardData';
import {
  loadOpportunityImport,
  normalizeDashboardDate,
  OPPORTUNITY_IMPORT_EVENT,
  normalizeBusinessValue,
} from '../utils/opportunityImport';

const STATIC_CLASSEUR_URL = '/data/Classeur.xlsx';

const sum = (arr, selector) => arr.reduce((acc, item) => acc + selector(item), 0);

const monthToDate = {
  Jan: '2026-01-01',
  Feb: '2026-02-01',
  Mar: '2026-03-01',
  Apr: '2026-04-01',
  May: '2026-05-01',
  Jun: '2026-06-01',
};

const minFilterDate = '2026-01-01';
const maxFilterDate = '2026-06-30';

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

const QUOTATION_TAT_GROUPS = ['Sales', 'UW', 'Reinsurance', 'Risk engineer', 'Doctor'];
const POLICY_TAT_GROUPS = ['AML', 'Finance', 'Policy Admin'];

function groupBy(recordsList, key, valueSelector) {
  return Object.entries(
    recordsList.reduce((acc, item) => {
      const group = item[key];
      acc[group] = (acc[group] || 0) + valueSelector(item);
      return acc;
    }, {})
  ).map(([name, value]) => ({ name, value }));
}

const isValidIsoDate = (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value ?? ''));

const normalizeDashboardBusiness = (value) => {
  const normalized = normalizeBusinessValue(value);

  if (normalized === 'motor') return 'Motor';
  if (normalized === 'medical') return 'Medical';
  if (normalized === 'life') return 'Life';
  if (normalized === 'general') return 'General';

  return '';
};

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

const parseNumber = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  const normalized = String(value).replace(/,/g, '').trim();
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeStatus = (value) => String(value ?? '').trim().toLowerCase();

const normalizeAssignToGroup = (value) => {
  const raw = String(value ?? '').trim().toLowerCase();

  if (!raw) return '';

  if (raw === 'sales') return 'Sales';
  if (raw === 'crc') return 'CRC';
  if (raw === 'uw') return 'UW';
  if (raw === 'reinsurance') return 'Reinsurance';
  if (raw === 'risk engineer' || raw === 'riskengineer') return 'Risk engineer';
  if (raw === 'doctor') return 'Doctor';
  if (raw === 'aml') return 'AML';
  if (raw === 'finance') return 'Finance';
  if (raw === 'policy admin' || raw === 'policyadmin') return 'Policy Admin';

  return String(value ?? '').trim();
};

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

const normalizeClasseurDate = (value) => {
  const rawValue = String(value ?? '').trim();
  if (!rawValue) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(rawValue)) {
    return rawValue;
  }

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

const buildQuotationTatDataset = (rows, groups) => {
  return groups.map((groupName) => {
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
        : rowsWithRequestId.reduce((acc, row) => acc + parseNumber(row.UsrTurnaroundDays), 0) / count;

    const avgMinutes =
      count === 0
        ? 0
        : rowsWithRequestId.reduce((acc, row) => acc + parseNumber(row.UsrTurnaroundTime), 0) / count;

    return {
      name: groupName,
      count,
      avgDays: Number(avgDays.toFixed(2)),
      avgMinutes: Number(avgMinutes.toFixed(2)),
    };
  });
};

const buildPolicyTatDataset = (rows, groups) => {
  return groups.map((groupName) => {
    const groupRows = rows.filter(
      (row) => normalizeAssignToGroup(row.UsrAssignToGroup) === groupName
    );

    const count = groupRows.length;

    const avgDays =
      count === 0
        ? 0
        : groupRows.reduce((acc, row) => acc + parseNumber(row.UsrTurnaroundDays), 0) / count;

    const avgMinutes =
      count === 0
        ? 0
        : groupRows.reduce((acc, row) => acc + parseNumber(row.UsrTurnaroundTime), 0) / count;

    return {
      name: groupName,
      count,
      avgDays: Number(avgDays.toFixed(2)),
      avgMinutes: Number(avgMinutes.toFixed(2)),
    };
  });
};

const TatTooltip = ({ active, payload, label }) => {
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
        boxShadow: '0 8px 24px rgba(15,23,42,0.08)',
      }}
    >
      <div style={{ fontWeight: 700, marginBottom: 6 }}>{label}</div>
      <div>Count: {data.count}</div>
      <div>Avg TAT Days: {data.avgDays}</div>
      <div>Avg TAT Minutes: {data.avgMinutes}</div>
    </div>
  );
};

export default function DashboardPage() {
  const [importedWorkbook, setImportedWorkbook] = useState(() => loadOpportunityImport());
  const [classeurRows, setClasseurRows] = useState([]);
  const [classeurError, setClasseurError] = useState('');

  const importedSheet =
    importedWorkbook?.workbookData?.sheets?.find(
      (sheet) => sheet.name === importedWorkbook?.activeSheet
    ) ??
    importedWorkbook?.workbookData?.sheets?.[0] ??
    null;

  const importedRows = importedSheet?.rows ?? [];

  const importedFilterDates = importedRows
    .map((row) => normalizeDashboardDate(row.UsrQuoteSubmissionDate))
    .filter((value) => isValidIsoDate(value))
    .sort();

  const effectiveMinDate = importedFilterDates[0] ?? minFilterDate;
  const effectiveMaxDate = importedFilterDates[importedFilterDates.length - 1] ?? maxFilterDate;

  const [filters, setFilters] = useState({
    fromDate: effectiveMinDate,
    toDate: effectiveMaxDate,
    businesses: [],
  });

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
    const loadClasseurWorkbook = async () => {
      try {
        setClasseurError('');

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
        const firstSheet = workbook.Sheets[firstSheetName];

        const rawRows = XLSX.utils.sheet_to_json(firstSheet, {
          defval: '',
          raw: false,
        });

        setClasseurRows(rawRows);
      } catch (error) {
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

  const quotationsTatByDepartment = useMemo(() => {
    return buildQuotationTatDataset(filteredClasseurRows, QUOTATION_TAT_GROUPS);
  }, [filteredClasseurRows]);

  const policiesTatByDepartment = useMemo(() => {
    return buildPolicyTatDataset(filteredClasseurRows, POLICY_TAT_GROUPS);
  }, [filteredClasseurRows]);

  const quotationsCount = useMemo(() => filteredImportedRows.length, [filteredImportedRows]);

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

  const filteredRecords = useMemo(() => {
    return records.filter((record) => {
      const recordDate = monthToDate[record.month];
      const matchesDate = recordDate >= filters.fromDate && recordDate <= filters.toDate;

      const matchesBusiness =
        !Array.isArray(filters.businesses) ||
        filters.businesses.length === 0 ||
        filters.businesses.includes(record.lob);

      return matchesDate && matchesBusiness;
    });
  }, [filters.businesses, filters.fromDate, filters.toDate]);

  const quotesByStatus = useMemo(() => {
    return groupBy(
      filteredRecords.filter((r) => r.status !== 'New'),
      'status',
      (r) => r.quotations
    );
  }, [filteredRecords]);

  const conversionByLob = useMemo(() => {
    return Object.values(
      filteredRecords.reduce((acc, item) => {
        if (!acc[item.lob]) {
          acc[item.lob] = { name: item.lob, quotations: 0, policies: 0 };
        }
        acc[item.lob].quotations += item.quotations;
        acc[item.lob].policies += item.convertedPolicies;
        return acc;
      }, {})
    ).map((item) => ({
      name: item.name,
      value:
        item.quotations === 0
          ? 0
          : Number(((item.policies / item.quotations) * 100).toFixed(1)),
    }));
  }, [filteredRecords]);

  const durationBuckets = useMemo(() => {
    return groupBy(filteredRecords, 'durationBucket', (r) => r.convertedPolicies);
  }, [filteredRecords]);

  const gwpComparisonByRegion = useMemo(() => {
    return Object.values(
      filteredRecords.reduce((acc, item) => {
        if (!acc[item.region]) {
          acc[item.region] = { name: item.region, actual: 0, target: 0 };
        }
        acc[item.region].actual += item.actualGwp;
        acc[item.region].target += item.expectedGwp;
        return acc;
      }, {})
    );
  }, [filteredRecords]);

  const gwpComparisonByBusiness = useMemo(() => {
    return Object.values(
      filteredRecords.reduce((acc, item) => {
        if (!acc[item.lob]) {
          acc[item.lob] = { name: item.lob, actual: 0, target: 0 };
        }
        acc[item.lob].actual += item.actualGwp;
        acc[item.lob].target += item.expectedGwp;
        return acc;
      }, {})
    );
  }, [filteredRecords]);

  const gwpComparisonBySource = useMemo(() => {
    return Object.values(
      filteredRecords.reduce((acc, item) => {
        if (!acc[item.source]) {
          acc[item.source] = { name: item.source, actual: 0, target: 0 };
        }
        acc[item.source].actual += item.actualGwp;
        acc[item.source].target += item.expectedGwp;
        return acc;
      }, {})
    );
  }, [filteredRecords]);

  const pipeline = useMemo(() => {
    const leads = Math.round(metrics.quotations * 1.18);
    const quotations = metrics.quotations;

    const uw = sum(
      filteredRecords.filter((r) =>
        ['Pending UW', 'Approved', 'Converted', 'Rejected', 'Returned'].includes(r.status)
      ),
      (r) => r.quotations
    );

    const approved = sum(
      filteredRecords.filter((r) => ['Approved', 'Converted'].includes(r.status)),
      (r) => r.quotations
    );

    const issued = metrics.policies;

    return [
      { name: 'Leads', value: leads },
      { name: 'Quotations', value: quotations },
      { name: 'UW Review', value: uw },
      { name: 'Approved', value: approved },
      { name: 'Issued', value: issued },
    ];
  }, [filteredRecords, metrics]);

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
        <ChartCard title="Pipeline Overview">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={pipeline}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Tooltip />
              <Bar dataKey="value" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Quotations by Status">
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={quotesByStatus}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Tooltip />
              <Bar dataKey="value" radius={[8, 8, 0, 0]} />
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

      {classeurError ? (
        <section className="two-col">
          <div className="card">
            <strong>{classeurError}</strong>
          </div>
        </section>
      ) : null}

      <section className="three-col">
        <ChartCard title="Conversion Rate by Line of Business">
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={conversionByLob}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis unit="%" />
              <Tooltip />
              <Bar dataKey="value" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Policies Converted by Duration">
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={durationBuckets}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Tooltip />
              <Bar dataKey="value" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>

      <section className="two-col">
        <ChartCard title="GWP by Region: Actual vs Target">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={gwpComparisonByRegion} barGap={10}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Legend />
              <Tooltip
                formatter={(value, name) => [
                  `SAR ${value.toLocaleString()}`,
                  name === 'actual' ? 'Actual GWP' : 'Target GWP',
                ]}
              />
              <Bar dataKey="actual" name="Actual GWP" radius={[8, 8, 0, 0]} />
              <Bar dataKey="target" name="Target GWP" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>

      <section className="two-col">
        <ChartCard title="GWP by Business: Actual vs Target">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={gwpComparisonByBusiness} barGap={10}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Legend />
              <Tooltip
                formatter={(value, name) => [
                  `SAR ${value.toLocaleString()}`,
                  name === 'actual' ? 'Actual GWP' : 'Target GWP',
                ]}
              />
              <Bar dataKey="actual" name="Actual GWP" radius={[8, 8, 0, 0]} />
              <Bar dataKey="target" name="Target GWP" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="GWP by Source: Actual vs Target">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={gwpComparisonBySource} barGap={10}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Legend />
              <Tooltip
                formatter={(value, name) => [
                  `SAR ${value.toLocaleString()}`,
                  name === 'actual' ? 'Actual GWP' : 'Target GWP',
                ]}
              />
              <Bar dataKey="actual" name="Actual GWP" radius={[8, 8, 0, 0]} />
              <Bar dataKey="target" name="Target GWP" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>
    </>
  );
}
