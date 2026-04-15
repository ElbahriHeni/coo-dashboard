import { useEffect, useId, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  loadOpportunityImport,
  normalizeDashboardDate,
  normalizeHeader,
  OPPORTUNITY_COLUMNS,
  saveOpportunityImport,
  getMappedBusinessFromUsrClass,
} from '../utils/opportunityImport';

const PAGE_SIZE = 200;
const STATIC_WORKBOOK_URL = '/data/AllOppQDR.xlsx';
const BUSINESS_MAPPING_COLUMN = 'Business Mapping';

const formatFileSize = (size) => {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
};

const formatCellValue = (value) => {
  if (value === null || value === undefined || value === '') return '-';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
};

const formatOpportunityCellValue = (header, value) => {
  if (header === 'UsrQuoteSubmissionDate') {
    return formatCellValue(normalizeDashboardDate(value));
  }

  return formatCellValue(value);
};

const getExportRows = (rows) =>
  rows.map((row) => {
    const exportRow = OPPORTUNITY_COLUMNS.reduce((acc, header) => {
      acc[header] =
        header === 'UsrQuoteSubmissionDate' ? normalizeDashboardDate(row[header]) : row[header] ?? '';
      return acc;
    }, {});

    exportRow[BUSINESS_MAPPING_COLUMN] = row[BUSINESS_MAPPING_COLUMN] ?? '';
    return exportRow;
  });

const buildWorkbookData = (workbook, fileName, fileSize) => {
  const sheets = workbook.SheetNames.map((name) => {
    const worksheet = workbook.Sheets[name];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, {
      defval: '',
      raw: false,
    });

    const headerMap = Object.fromEntries(
      Object.keys(rawRows[0] ?? {}).map((header) => [normalizeHeader(header), header])
    );

    const matchedColumns = OPPORTUNITY_COLUMNS.filter(
      (column) => headerMap[normalizeHeader(column)]
    );

    const rows = rawRows.map((row) => {
      const mappedRow = OPPORTUNITY_COLUMNS.reduce((acc, column) => {
        const sourceKey = headerMap[normalizeHeader(column)];
        const value = sourceKey ? row[sourceKey] : '';
        acc[column] =
          column === 'UsrQuoteSubmissionDate' ? normalizeDashboardDate(value) : value;
        return acc;
      }, {});

      mappedRow[BUSINESS_MAPPING_COLUMN] = getMappedBusinessFromUsrClass(mappedRow.UsrClass).business;

      return mappedRow;
    });

    return {
      name,
      rowCount: rows.length,
      matchedColumns,
      rows,
    };
  });

  return {
    fileName,
    fileSize,
    sheetCount: sheets.length,
    sheets,
  };
};

export default function SpreadsheetPage() {
  const inputId = useId();
  const inputRef = useRef(null);
  const [isDragging, setIsDragging] = useState(false);
  const [uploadError, setUploadError] = useState('');
  const [isAutoLoading, setIsAutoLoading] = useState(false);
  const [workbookData, setWorkbookData] = useState(() => loadOpportunityImport()?.workbookData ?? null);
  const [activeSheet, setActiveSheet] = useState(() => loadOpportunityImport()?.activeSheet ?? '');
  const [currentPage, setCurrentPage] = useState(1);

  const sheetPreview = useMemo(() => {
    if (!workbookData || !activeSheet) return null;
    return workbookData.sheets.find((sheet) => sheet.name === activeSheet) ?? null;
  }, [activeSheet, workbookData]);

  const totalRows = sheetPreview?.rows.length ?? 0;
  const totalPages = Math.max(1, Math.ceil(totalRows / PAGE_SIZE));

  const visibleRows = useMemo(() => {
    if (!sheetPreview) return [];
    const startIndex = (currentPage - 1) * PAGE_SIZE;
    return sheetPreview.rows.slice(startIndex, startIndex + PAGE_SIZE);
  }, [currentPage, sheetPreview]);

  const startRow = totalRows === 0 ? 0 : (currentPage - 1) * PAGE_SIZE + 1;
  const endRow = Math.min(currentPage * PAGE_SIZE, totalRows);

  useEffect(() => {
    setCurrentPage(1);
  }, [activeSheet, workbookData]);

  useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages);
    }
  }, [currentPage, totalPages]);

  useEffect(() => {
    if (!workbookData) return;

    saveOpportunityImport({
      workbookData,
      activeSheet,
    });
  }, [activeSheet, workbookData]);

  useEffect(() => {
    if (workbookData) return;

    const autoLoadWorkbook = async () => {
      try {
        setIsAutoLoading(true);
        setUploadError('');

        const response = await fetch(STATIC_WORKBOOK_URL, { cache: 'no-store' });
        if (!response.ok) {
          throw new Error('Static workbook not found');
        }

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        const builtWorkbookData = buildWorkbookData(workbook, 'AllOppQDR.xlsx', buffer.byteLength);
        const firstUsableSheet =
          builtWorkbookData.sheets.find((sheet) => sheet.matchedColumns.length > 0) ??
          builtWorkbookData.sheets[0] ??
          null;

        setWorkbookData(builtWorkbookData);
        setActiveSheet(firstUsableSheet?.name ?? '');
      } catch (error) {
        setUploadError('The deployed workbook could not be loaded automatically.');
      } finally {
        setIsAutoLoading(false);
      }
    };

    autoLoadWorkbook();
  }, [workbookData]);

  const loadWorkbook = async (file) => {
    if (!file) return;

    const isSpreadsheet =
      file.name.toLowerCase().endsWith('.xlsx') ||
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    if (!isSpreadsheet) {
      setUploadError('Please drop an .xlsx file so I can load the opportunity table.');
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const builtWorkbookData = buildWorkbookData(workbook, file.name, file.size);

      const firstUsableSheet =
        builtWorkbookData.sheets.find((sheet) => sheet.matchedColumns.length > 0) ??
        builtWorkbookData.sheets[0] ??
        null;

      setWorkbookData(builtWorkbookData);
      setActiveSheet(firstUsableSheet?.name ?? '');
      setUploadError('');
    } catch (error) {
      setUploadError('That file could not be read as a workbook. Please try another .xlsx file.');
      setWorkbookData(null);
      setActiveSheet('');
    }
  };

  const onInputChange = async (event) => {
    const [file] = event.target.files ?? [];
    await loadWorkbook(file);
    event.target.value = '';
  };

  const onDrop = async (event) => {
    event.preventDefault();
    setIsDragging(false);
    const [file] = event.dataTransfer.files ?? [];
    await loadWorkbook(file);
  };

  const exportActiveSheet = () => {
    if (!sheetPreview) return;

    const exportWorkbook = XLSX.utils.book_new();
    const exportSheet = XLSX.utils.json_to_sheet(getExportRows(sheetPreview.rows), {
      header: [...OPPORTUNITY_COLUMNS, BUSINESS_MAPPING_COLUMN],
    });

    XLSX.utils.book_append_sheet(exportWorkbook, exportSheet, sheetPreview.name || 'Opportunities');

    const baseName = workbookData?.fileName?.replace(/\.xlsx$/i, '') ?? 'opportunities';
    const safeSheetName = String(sheetPreview.name || 'sheet').replace(/[\\/:*?"<>|]/g, '_');
    XLSX.writeFile(exportWorkbook, `${baseName}-${safeSheetName}-normalized.xlsx`);
  };

  return (
    <section className="spreadsheet-layout">
      <div
        className={`card uploader-card${isDragging ? ' dragging' : ''}`}
        onDragEnter={(event) => {
          event.preventDefault();
          setIsDragging(true);
        }}
        onDragOver={(event) => {
          event.preventDefault();
          setIsDragging(true);
        }}
        onDragLeave={(event) => {
          event.preventDefault();
          if (event.currentTarget === event.target) {
            setIsDragging(false);
          }
        }}
        onDrop={onDrop}
      >
        <div className="uploader-copy">
          <p className="eyebrow uploader-eyebrow">Excel Import</p>
          <h2>Load opportunities from XLSX</h2>
          <p className="subtitle uploader-subtitle">
            This page auto-loads the deployed workbook, maps it into the dashboard fields, and stores
            the computed business mapping for downstream dashboard filters.
          </p>
        </div>

        <div className="uploader-actions">
          <input
            ref={inputRef}
            id={inputId}
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            className="sr-only"
            onChange={onInputChange}
          />
          <label htmlFor={inputId} className="primary-button">
            Replace with another XLSX
          </label>
          <button
            type="button"
            className="secondary-button"
            onClick={() => inputRef.current?.click()}
          >
            Browse Again
          </button>
        </div>

        {isAutoLoading ? <p className="meta-label">Loading deployed workbook...</p> : null}
        {uploadError ? <p className="upload-error">{uploadError}</p> : null}
      </div>

      <div className="spreadsheet-meta-grid">
        <article className="card meta-card">
          <span className="meta-label">File</span>
          <strong>{workbookData?.fileName ?? 'No workbook loaded yet'}</strong>
        </article>
        <article className="card meta-card">
          <span className="meta-label">Size</span>
          <strong>{workbookData ? formatFileSize(workbookData.fileSize) : '-'}</strong>
        </article>
        <article className="card meta-card">
          <span className="meta-label">Sheets</span>
          <strong>{workbookData?.sheetCount ?? 0}</strong>
        </article>
        <article className="card meta-card">
          <span className="meta-label">Imported Rows</span>
          <strong>{sheetPreview?.rows.length ?? 0}</strong>
        </article>
      </div>

      {workbookData ? (
        <div className="card spreadsheet-card">
          <div className="sheet-tabs" role="tablist" aria-label="Workbook sheets">
            {workbookData.sheets.map((sheet) => (
              <button
                key={sheet.name}
                type="button"
                role="tab"
                aria-selected={sheet.name === activeSheet}
                className={`sheet-tab${sheet.name === activeSheet ? ' active' : ''}`}
                onClick={() => setActiveSheet(sheet.name)}
              >
                <span>{sheet.name}</span>
                <small>
                  {sheet.rowCount} rows / {sheet.matchedColumns.length} matched columns
                </small>
              </button>
            ))}
          </div>

          <div className="sheet-summary">
            <div>
              <h3>{sheetPreview?.name}</h3>
              <p>
                Showing rows {startRow}-{endRow} of {totalRows} imported opportunity rows.
              </p>
            </div>
            {totalRows > 0 ? (
              <div className="uploader-actions">
                <button
                  type="button"
                  className="primary-button"
                  onClick={exportActiveSheet}
                >
                  Export Normalized XLSX
                </button>
                <button
                  type="button"
                  className="secondary-button"
                  onClick={() => setCurrentPage((page) => Math.max(1, page - 1))}
                  disabled={currentPage === 1}
                >
                  Previous
                </button>
                <span className="meta-label">
                  Page {currentPage} of {totalPages}
                </span>
                <button
                  type="button"
                  className="secondary-button"
                  onClick={() => setCurrentPage((page) => Math.min(totalPages, page + 1))}
                  disabled={currentPage === totalPages}
                >
                  Next
                </button>
              </div>
            ) : null}
          </div>

          {sheetPreview && sheetPreview.matchedColumns.length === 0 ? (
            <p className="upload-error">
              This sheet does not contain the expected `Usr...` headers yet. Select another sheet or
              replace the workbook with one that contains those columns.
            </p>
          ) : null}

          <div className="table-scroll spreadsheet-table-wrap">
            <table>
              <thead>
                <tr>
                  {OPPORTUNITY_COLUMNS.map((header) => (
                    <th key={header}>{header}</th>
                  ))}
                  <th>{BUSINESS_MAPPING_COLUMN}</th>
                </tr>
              </thead>
              <tbody>
                {visibleRows.map((row, rowIndex) => {
                  const mapping = getMappedBusinessFromUsrClass(row.UsrClass);

                  return (
                    <tr key={`${sheetPreview?.name}-${startRow + rowIndex}`}>
                      {OPPORTUNITY_COLUMNS.map((header) => (
                        <td key={`${header}-${startRow + rowIndex}`}>
                          {formatOpportunityCellValue(header, row[header])}
                        </td>
                      ))}
                      <td key={`${BUSINESS_MAPPING_COLUMN}-${startRow + rowIndex}`}>
                        {row[BUSINESS_MAPPING_COLUMN] || mapping.display}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="card empty-state-card">
          <h3>Ready for your workbook</h3>
          <p>
            The page will try to load the deployed workbook automatically. You can still replace it
            with another file if needed.
          </p>
        </div>
      )}
    </section>
  );
}
