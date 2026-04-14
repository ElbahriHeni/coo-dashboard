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

export const normalizeDashboardDate = (value) => {
  const rawValue = String(value ?? '').trim();
  if (!rawValue) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(rawValue)) return rawValue;

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
