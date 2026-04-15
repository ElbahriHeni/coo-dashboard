import React from 'react';

const BUSINESS_OPTIONS = ['General', 'Motor', 'Medical', 'Life'];

export default function FilterBar({
  filters,
  onFilterChange,
  regions = [],
  sources = [],
  departments = [],
}) {
  const selectedBusinesses = Array.isArray(filters.businesses) ? filters.businesses : [];

  const handleInputChange = (event) => {
    const { name, value } = event.target;
    onFilterChange(name, value);
  };

  const handleBusinessChange = (event) => {
    const values = Array.from(event.target.selectedOptions, (option) => option.value);
    onFilterChange('businesses', values);
  };

  return (
    <div className="card filter-bar">
      <div className="filter-grid">
        <div className="filter-field">
          <label htmlFor="fromDate">Date From</label>
          <input
            id="fromDate"
            name="fromDate"
            type="date"
            value={filters.fromDate || ''}
            onChange={handleInputChange}
          />
        </div>

        <div className="filter-field">
          <label htmlFor="toDate">Date To</label>
          <input
            id="toDate"
            name="toDate"
            type="date"
            value={filters.toDate || ''}
            onChange={handleInputChange}
          />
        </div>

        <div className="filter-field">
          <label htmlFor="businesses">Business</label>
          <select
            id="businesses"
            name="businesses"
            multiple
            value={selectedBusinesses}
            onChange={handleBusinessChange}
            className="multi-select"
          >
            {BUSINESS_OPTIONS.map((business) => (
              <option key={business} value={business}>
                {business}
              </option>
            ))}
          </select>
          <small className="filter-help">Hold Ctrl or Cmd to select multiple</small>
        </div>

        <div className="filter-field">
          <label htmlFor="region">Region</label>
          <select
            id="region"
            name="region"
            value={filters.region || 'All'}
            onChange={handleInputChange}
          >
            <option value="All">All</option>
            {regions.map((region) => (
              <option key={region} value={region}>
                {region}
              </option>
            ))}
          </select>
        </div>

        <div className="filter-field">
          <label htmlFor="source">Source</label>
          <select
            id="source"
            name="source"
            value={filters.source || 'All'}
            onChange={handleInputChange}
          >
            <option value="All">All</option>
            {sources.map((source) => (
              <option key={source} value={source}>
                {source}
              </option>
            ))}
          </select>
        </div>

        <div className="filter-field">
          <label htmlFor="department">Department</label>
          <select
            id="department"
            name="department"
            value={filters.department || 'All'}
            onChange={handleInputChange}
          >
            <option value="All">All</option>
            {departments.map((department) => (
              <option key={department} value={department}>
                {department}
              </option>
            ))}
          </select>
        </div>
      </div>
    </div>
  );
}
