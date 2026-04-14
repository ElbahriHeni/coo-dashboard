import { useEffect, useMemo, useState } from 'react';
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
import { records, targets } from '../data/dashboardData';
import {
  loadOpportunityImport,
  normalizeDashboardDate,
  OPPORTUNITY_IMPORT_EVENT,
} from '../utils/opportunityImport';

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

function groupBy(recordsList, key, valueSelector) {
  return Object.entries(
    recordsList.reduce((acc, item) => {
      const group = item[key];
      acc[group] = (acc[group] || 0) + valueSelector(item);
      return acc;
    }, {})
  ).map(([name, value]) => ({ name, value }));
}

export default function DashboardPage() {
  const [importedWorkbook, setImportedWorkbook] = useState(() => loadOpportunityImport());
  const importedSheet =
    importedWorkbook?.workbookData?.sheets?.find((sheet) => sheet.name === importedWorkbook.activeSheet) ??
    importedWorkbook?.workbookData?.sheets?.[0] ??
    null;
  const importedRows = importedSheet?.rows ?? [];
  const importedFilterDates = importedRows
    .map((row) => normalizeDashboardDate(row.UsrQuoteSubmissionDate))
    .filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value ?? '')))
    .sort();
  const effectiveMinDate = importedFilterDates[0] ?? minFilterDate;
  const effectiveMaxDate = importedFilterDates[importedFilterDates.length - 1] ?? maxFilterDate;

  const [filters, setFilters] = useState({
    from: effectiveMinDate,
    to: effectiveMaxDate,
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
    setFilters((prev) => {
      const nextFrom =
        prev.from < effectiveMinDate || prev.from > effectiveMaxDate ? effectiveMinDate : prev.from;
      const nextTo =
        prev.to < effectiveMinDate || prev.to > effectiveMaxDate ? effectiveMaxDate : prev.to;

      if (nextFrom > nextTo) {
        return {
          from: effectiveMinDate,
          to: effectiveMaxDate,
        };
      }

      if (nextFrom === prev.from && nextTo === prev.to) {
        return prev;
      }

      return {
        from: nextFrom,
        to: nextTo,
      };
    });
  }, [effectiveMaxDate, effectiveMinDate]);

  const filteredImportedRows = useMemo(() => {
    if (importedRows.length === 0) return [];

    const isFullImportedRange =
      filters.from === effectiveMinDate && filters.to === effectiveMaxDate;

    if (isFullImportedRange) {
      return importedRows;
    }

    return importedRows.filter((row) => {
      const submissionDate = normalizeDashboardDate(row.UsrQuoteSubmissionDate);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(submissionDate ?? ''))) {
        return false;
      }
      return submissionDate >= filters.from && submissionDate <= filters.to;
    });
  }, [effectiveMaxDate, effectiveMinDate, filters.from, filters.to, importedRows]);

  const filteredRecords = useMemo(() => {
    return records.filter((record) => {
      const recordDate = monthToDate[record.month];
      return recordDate >= filters.from && recordDate <= filters.to;
    });
  }, [filters]);

  const metrics = useMemo(() => {
    const quotations =
      importedRows.length > 0
        ? filteredImportedRows.length
        : sum(filteredRecords, (r) => r.quotations);
    const policies = sum(filteredRecords, (r) => r.convertedPolicies);
    const conversionRate = quotations === 0 ? 0 : (policies / quotations) * 100;
    const expectedGwpMotor = sum(
      filteredRecords.filter((r) => r.lob === 'Motor'),
      (r) => r.expectedGwp
    );

    return { quotations, policies, conversionRate, expectedGwpMotor };
  }, [filteredImportedRows.length, filteredRecords, importedRows.length]);

  const quotesByStatus = useMemo(() => {
    return groupBy(
      filteredRecords.filter((r) => r.status !== 'New'),
      'status',
      (r) => r.quotations
    );
  }, [filteredRecords]);

  const assignedByDepartment = useMemo(
    () => groupBy(filteredRecords, 'department', (r) => r.quotations),
    [filteredRecords]
  );

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
      value: item.quotations === 0 ? 0 : Number(((item.policies / item.quotations) * 100).toFixed(1)),
    }));
  }, [filteredRecords]);

  const durationBuckets = useMemo(
    () => groupBy(filteredRecords, 'durationBucket', (r) => r.convertedPolicies),
    [filteredRecords]
  );

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

  const onFilterChange = (key, value) => setFilters((prev) => ({ ...prev, [key]: value }));

  return (
    <>
      <FilterBar
        filters={filters}
        onChange={onFilterChange}
        minDate={effectiveMinDate}
        maxDate={effectiveMaxDate}
      />

      <section className="kpi-grid">
        <KpiCard
          title="Quotations"
          value={metrics.quotations}
          target={targets.quotations}
          showTarget={false}
        />
        <KpiCard
          title="Policies Converted"
          value={metrics.policies}
          target={targets.policies}
          showTarget={false}
        />
        <KpiCard
          title="Conversion Rate"
          value={metrics.conversionRate}
          target={targets.conversionRate}
          suffix="%"
          showTarget={false}
        />
        <KpiCard
          title="Expected GWP"
          value={metrics.expectedGwpMotor}
          target={targets.expectedGwpMotor}
          currency
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

      <section className="three-col">
        <ChartCard title="Assigned to Each Department">
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={assignedByDepartment} layout="vertical" margin={{ left: 30 }}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis type="number" />
              <YAxis type="category" dataKey="name" width={100} />
              <Tooltip />
              <Bar dataKey="value" radius={[0, 8, 8, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

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
              <Bar dataKey="actual" name="Actual GWP" fill="#1d4ed8" radius={[8, 8, 0, 0]} />
              <Bar dataKey="target" name="Target GWP" fill="#94a3b8" radius={[8, 8, 0, 0]} />
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
              <Bar dataKey="actual" name="Actual GWP" fill="#1d4ed8" radius={[8, 8, 0, 0]} />
              <Bar dataKey="target" name="Target GWP" fill="#94a3b8" radius={[8, 8, 0, 0]} />
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
              <Bar dataKey="actual" name="Actual GWP" fill="#1d4ed8" radius={[8, 8, 0, 0]} />
              <Bar dataKey="target" name="Target GWP" fill="#94a3b8" radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </section>
    </>
  );
}
