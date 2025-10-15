// src/components/ui/Badge.jsx
import React from 'react';
import './Badge.css';

/**
 * Badge Component
 * @param {string} children - Text to display inside the badge
 * @param {string} variant - One of: 'primary', 'success', 'warning', 'error', 'info',
 *                           'principal', 'teacher', 'dos', 'neutral'
 * @param {string} size - 'sm' | 'md' (default: 'md')
 * @param {string} className - Additional custom classes
 */
const Badge = ({ 
  children, 
  variant = 'neutral', 
  size = 'md', 
  className = '' 
}) => {
  const baseClass = 'badge';
  const variantClass = `badge--${variant}`;
  const sizeClass = `badge--${size}`;
  
  return (
    <span className={`${baseClass} ${variantClass} ${sizeClass} ${className}`}>
      {children}
    </span>
  );
};

export default Badge;