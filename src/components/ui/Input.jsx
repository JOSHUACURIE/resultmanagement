// src/components/ui/Input.jsx
import React from 'react';
import './Input.css';

/**
 * Input Component
 * @param {string} label - Optional label text
 * @param {string} id - Required for accessibility (matches label htmlFor)
 * @param {string} type - Input type (text, email, number, password, etc.)
 * @param {string} placeholder - Placeholder text
 * @param {boolean} required - Adds required asterisk and HTML required attribute
 * @param {string} error - Error message to display
 * @param {string} className - Additional custom classes
 * @param {React.ReactNode} startIcon - Optional icon on the left (e.g., 📧)
 * @param {React.ReactNode} endIcon - Optional icon on the right (e.g., 👁️)
 * @param {object} rest - All other props (value, onChange, disabled, etc.)
 */
const Input = ({
  label,
  id,
  type = 'text',
  placeholder,
  required = false,
  error = null,
  className = '',
  startIcon = null,
  endIcon = null,
  ...rest
}) => {
  return (
    <div className={`input-group ${error ? 'input-group--error' : ''} ${className}`}>
      {label && (
        <label htmlFor={id} className="input-label">
          {label}
          {required && <span className="input-required">*</span>}
        </label>
      )}

      <div className="input-wrapper">
        {startIcon && <span className="input-icon input-icon--start">{startIcon}</span>}
        
        <input
          id={id}
          type={type}
          placeholder={placeholder}
          className={`input-field ${startIcon ? 'input-field--with-start' : ''} ${
            endIcon ? 'input-field--with-end' : ''
          }`}
          aria-invalid={!!error}
          aria-describedby={error ? `${id}-error` : undefined}
          required={required}
          {...rest}
        />
        
        {endIcon && <span className="input-icon input-icon--end">{endIcon}</span>}
      </div>

      {error && (
        <p id={`${id}-error`} className="input-error">
          {error}
        </p>
      )}
    </div>
  );
};

export default Input;