import React from 'react';
import './Table.css';

/**
 * Table Component (Fully Responsive)
 * @param {Array} columns - Array of column config: { key, header, render (optional) }
 * @param {Array|Object} data - Array of row objects OR API response object
 * @param {string} emptyMessage - Message when no data (default: "No records found")
 * @param {boolean} loading - Show loading state
 * @param {string} className - Additional classes
 * @param {string} dataKey - Key to extract array from response object (default: 'data')
 */
const Table = ({ 
  columns, 
  data, 
  emptyMessage = "No records found", 
  loading = false,
  className = '',
  dataKey = 'data'
}) => {
  const tableData = React.useMemo(() => {
    if (!data) return [];
    if (Array.isArray(data)) return data;
    if (data && typeof data === 'object' && data[dataKey]) {
      return Array.isArray(data[dataKey]) ? data[dataKey] : [];
    }
    console.warn('⚠️ Table component: data is not an array. Expected array or object with data property.', data);
    return [];
  }, [data, dataKey]);

  // 🕓 Loading state
  if (loading) {
    return (
      <div className={`table-container ${className}`}>
        <table className="data-table">
          <tbody>
            <tr>
              <td colSpan={columns?.length || 1} className="table-loading">Loading...</td>
            </tr>
          </tbody>
        </table>
      </div>
    );
  }

  // 📭 Empty state
  if (!tableData || tableData.length === 0) {
    return (
      <div className={`table-container ${className}`}>
        <table className="data-table">
          <tbody>
            <tr>
              <td colSpan={columns?.length || 1} className="table-empty">{emptyMessage}</td>
            </tr>
          </tbody>
        </table>
      </div>
    );
  }

  // ❌ No columns defined
  if (!columns || !Array.isArray(columns) || columns.length === 0) {
    return (
      <div className={`table-container ${className}`}>
        <table className="data-table">
          <tbody>
            <tr>
              <td className="table-error">No columns defined for table</td>
            </tr>
          </tbody>
        </table>
      </div>
    );
  }

  // ✅ Render actual data
  return (
    <div className={`table-container ${className}`}>
      <table className="data-table">
        <thead>
          <tr>
            {columns.map((col) => (
              <th key={col.key} className="table-header">
                {col.header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {tableData.map((row, index) => (
            <tr key={row.id || index} className="table-row">
              {columns.map((col) => (
                <td
                  key={col.key}
                  className="table-cell"
                  data-label={col.header} // 👈 Used in mobile view
                >
                  {col.render ? col.render(row[col.key], row) : row[col.key]}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default Table;
