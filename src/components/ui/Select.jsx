// src/components/ui/Select.jsx
import React from 'react';
import './Select.css';

/**
 * Select Component
 * @param {string} label - Label text for the select input
 * @param {string} name - Input name and id
 * @param {string|number} value - Current value
 * @param {function} onChange - Change event handler
 * @param {string} error - Error message to display
 * @param {boolean} required - Marks field as required
 * @param {React.ReactNode} children - Option elements
 */
const Select = ({ label, name, value, onChange, error, required, children }) => {
  return (
    <div className={`form-group ${error ? 'form-group--error' : ''}`}>
      {label && (
        <label htmlFor={name} className="form-label">
          {label} {required && <span className="required">*</span>}
        </label>
      )}

      <div className="select-wrapper">
        <select
          id={name}
          name={name}
          value={value}
          onChange={onChange}
          className={`form-select ${error ? 'input-error' : ''}`}
          required={required}
        >
          {children}
        </select>
      </div>

      {error && <div className="form-error">{error}</div>}
    </div>
  );
};

export default Select;
