import { useEffect, useRef, useState } from 'react';
import { formatDisplayDate, parseDisplayDate } from '../utils/opportunityImport';

export default function FilterBar({ filters, onChange, minDate, maxDate }) {
  const [fromInput, setFromInput] = useState(formatDisplayDate(filters.from));
  const [toInput, setToInput] = useState(formatDisplayDate(filters.to));
  const fromPickerRef = useRef(null);
  const toPickerRef = useRef(null);

  useEffect(() => {
    setFromInput(formatDisplayDate(filters.from));
  }, [filters.from]);

  useEffect(() => {
    setToInput(formatDisplayDate(filters.to));
  }, [filters.to]);

  const commitChange = (key, value) => {
    const parsedValue = parseDisplayDate(value);
    if (!parsedValue) return;

    if (key === 'from') {
      if (parsedValue < minDate || parsedValue > (filters.to || maxDate)) return;
    } else if (parsedValue < (filters.from || minDate) || parsedValue > maxDate) {
      return;
    }

    onChange(key, parsedValue);
  };

  return (
    <div className="card filters-grid">
      <label className="filter-item">
        <span>From</span>
        <div className="filter-date-input">
          <input
            type="text"
            inputMode="numeric"
            placeholder="DD-MM-YYYY"
            value={fromInput}
            onChange={(e) => {
              const nextValue = e.target.value;
              setFromInput(nextValue);
              commitChange('from', nextValue);
            }}
            onBlur={() => setFromInput(formatDisplayDate(filters.from))}
          />
          <button
            type="button"
            className="secondary-button"
            onClick={() => fromPickerRef.current?.showPicker?.()}
          >
            Pick
          </button>
          <input
            ref={fromPickerRef}
            type="date"
            className="sr-only"
            value={filters.from}
            min={minDate}
            max={filters.to || maxDate}
            onChange={(e) => onChange('from', e.target.value)}
            tabIndex={-1}
            aria-hidden="true"
          />
        </div>
      </label>
      <label className="filter-item">
        <span>To</span>
        <div className="filter-date-input">
          <input
            type="text"
            inputMode="numeric"
            placeholder="DD-MM-YYYY"
            value={toInput}
            onChange={(e) => {
              const nextValue = e.target.value;
              setToInput(nextValue);
              commitChange('to', nextValue);
            }}
            onBlur={() => setToInput(formatDisplayDate(filters.to))}
          />
          <button
            type="button"
            className="secondary-button"
            onClick={() => toPickerRef.current?.showPicker?.()}
          >
            Pick
          </button>
          <input
            ref={toPickerRef}
            type="date"
            className="sr-only"
            value={filters.to}
            min={filters.from || minDate}
            max={maxDate}
            onChange={(e) => onChange('to', e.target.value)}
            tabIndex={-1}
            aria-hidden="true"
          />
        </div>
      </label>
    </div>
  );
}
