export const OPPORTUNITY_COLUMNS = [
  'UsrOpportunityName',
  'UsrSource',
  'UsrRenewalDate',
  'UsrEstGWP',
  'UsrRegion',
  'UsrClient',
  'UsrSourceName',
  'UsrClientType',
  'UsrRiskCategory',
  'UsrClass',
  'UsrSubClass',
  'UsrLostReason',
  'UsrStatus',
  'UsrQuoteSubmissionDate',
  'UsrQuoteReceiptDate',
  'UsrRemarks',
];

const MONTH_INDEX = {
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

const STORAGE_KEY = 'crm-opportunity-import';
export const OPPORTUNITY_IMPORT_EVENT = 'opportunity-import-updated';

export const normalizeHeader = (value) => String(value ?? '').trim().toLowerCase();

export const normalizeText = (value) => String(value ?? '').trim();

export const normalizeTextLower = (value) => normalizeText(value).toLowerCase();

export const normalizeDashboardDate = (value) => {
  const rawValue = String(value ?? '').trim();
  if (!rawValue) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(rawValue)) {
    return rawValue;
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(rawValue)) {
    const [day, month, year] = rawValue.split('/');
    return `${year}-${month}-${day}`;
  }

  if (/^\d{2}-\d{2}-\d{4}$/.test(rawValue)) {
    const [day, month, year] = rawValue.split('-');
    return `${year}-${month}-${day}`;
  }

  const match = rawValue.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})(?:\s|$)/);
  if (!match) return rawValue;

  const [, day, monthToken, yearToken] = match;
  const month = MONTH_INDEX[monthToken.toUpperCase()];
  if (!month) return rawValue;

  const year = yearToken.length === 2 ? `20${yearToken}` : yearToken.padStart(4, '0');
  return `${year}-${month}-${day.padStart(2, '0')}`;
};

export const formatDisplayDate = (value) => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value ?? ''))) return value ?? '';

  const [year, month, day] = value.split('-');
  return `${day}-${month}-${year}`;
};

export const parseDisplayDate = (value) => {
  const trimmedValue = String(value ?? '').trim();
  if (!trimmedValue) return '';

  const match = trimmedValue.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!match) return null;

  const [, day, month, year] = match;
  const isoValue = `${year}-${month}-${day}`;
  const parsedDate = new Date(`${isoValue}T00:00:00`);

  if (Number.isNaN(parsedDate.getTime())) return null;
  if (parsedDate.toISOString().slice(0, 10) !== isoValue) return null;

  return isoValue;
};

export const saveOpportunityImport = (value) => {
  if (typeof window === 'undefined') return;

  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(value));
    window.dispatchEvent(new CustomEvent(OPPORTUNITY_IMPORT_EVENT, { detail: value }));
  } catch {
    // Ignore storage failures so uploads still work in-memory.
  }
};

export const loadOpportunityImport = () => {
  if (typeof window === 'undefined') return null;

  try {
    const rawValue = window.localStorage.getItem(STORAGE_KEY);
    return rawValue ? JSON.parse(rawValue) : null;
  } catch {
    return null;
  }
};

export const clearOpportunityImport = () => {
  if (typeof window === 'undefined') return;

  try {
    window.localStorage.removeItem(STORAGE_KEY);
    window.dispatchEvent(new CustomEvent(OPPORTUNITY_IMPORT_EVENT, { detail: null }));
  } catch {
    // Ignore storage failures.
  }
};

export const getActiveImportedSheet = () => {
  const importedState = loadOpportunityImport();
  const workbookData = importedState?.workbookData ?? null;
  const activeSheetName = importedState?.activeSheet ?? '';

  if (!workbookData?.sheets?.length) return null;

  return (
    workbookData.sheets.find((sheet) => sheet.name === activeSheetName) ??
    workbookData.sheets[0] ??
    null
  );
};

export const getImportedOpportunityRows = () => {
  const activeSheet = getActiveImportedSheet();
  return activeSheet?.rows ?? [];
};

const isAllFilter = (value) => {
  const normalized = normalizeTextLower(value);
  return !normalized || normalized === 'all';
};

const matchesTextFilter = (rowValue, filterValue) => {
  if (isAllFilter(filterValue)) return true;
  return normalizeTextLower(rowValue) === normalizeTextLower(filterValue);
};

const matchesDateRange = (rowDateValue, fromDate, toDate) => {
  const normalizedRowDate = normalizeDashboardDate(rowDateValue);

  if (!normalizedRowDate || !/^\d{4}-\d{2}-\d{2}$/.test(normalizedRowDate)) {
    return false;
  }

  const normalizedFromDate = normalizeDashboardDate(fromDate);
  const normalizedToDate = normalizeDashboardDate(toDate);

  if (normalizedFromDate && /^\d{4}-\d{2}-\d{2}$/.test(normalizedFromDate) && normalizedRowDate < normalizedFromDate) {
    return false;
  }

  if (normalizedToDate && /^\d{4}-\d{2}-\d{2}$/.test(normalizedToDate) && normalizedRowDate > normalizedToDate) {
    return false;
  }

  return true;
};

export const applyDashboardFiltersToRows = (rows, filters = {}) => {
  const safeRows = Array.isArray(rows) ? rows : [];

  const {
    region,
    source,
    sourceName,
    client,
    clientType,
    riskCategory,
    lineOfBusiness,
    className,
    subClass,
    status,
    lostReason,
    remarks,
    fromDate,
    toDate,
    quoteSubmissionDateFrom,
    quoteSubmissionDateTo,
  } = filters;

  const effectiveFromDate = quoteSubmissionDateFrom ?? fromDate ?? '';
  const effectiveToDate = quoteSubmissionDateTo ?? toDate ?? '';

  return safeRows.filter((row) => {
    const matchesRegion = matchesTextFilter(row.UsrRegion, region);
    const matchesSource = matchesTextFilter(row.UsrSource, source);
    const matchesSourceName = matchesTextFilter(row.UsrSourceName, sourceName);
    const matchesClient = matchesTextFilter(row.UsrClient, client);
    const matchesClientType = matchesTextFilter(row.UsrClientType, clientType);
    const matchesRiskCategory = matchesTextFilter(row.UsrRiskCategory, riskCategory);
    const matchesLineOfBusiness = matchesTextFilter(row.UsrClass, lineOfBusiness);
    const matchesClass = matchesTextFilter(row.UsrClass, className);
    const matchesSubClass = matchesTextFilter(row.UsrSubClass, subClass);
    const matchesStatus = matchesTextFilter(row.UsrStatus, status);
    const matchesLostReason = matchesTextFilter(row.UsrLostReason, lostReason);
    const matchesRemarks = matchesTextFilter(row.UsrRemarks, remarks);

    const matchesQuoteSubmissionDate =
      !effectiveFromDate && !effectiveToDate
        ? true
        : matchesDateRange(row.UsrQuoteSubmissionDate, effectiveFromDate, effectiveToDate);

    return (
      matchesRegion &&
      matchesSource &&
      matchesSourceName &&
      matchesClient &&
      matchesClientType &&
      matchesRiskCategory &&
      matchesLineOfBusiness &&
      matchesClass &&
      matchesSubClass &&
      matchesStatus &&
      matchesLostReason &&
      matchesRemarks &&
      matchesQuoteSubmissionDate
    );
  });
};

export const getFilteredImportedOpportunityRows = (filters = {}) => {
  const rows = getImportedOpportunityRows();
  return applyDashboardFiltersToRows(rows, filters);
};

export const getQuotationsCount = (filters = {}) => {
  return getFilteredImportedOpportunityRows(filters).length;
};
