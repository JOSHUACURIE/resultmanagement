
import React from 'react';
import './Card.css';

/**
 * Card Component
 * @param {React.ReactNode} children - Content inside the card
 * @param {string} className - Additional custom classes
 * @param {boolean} noPadding - Remove default padding (for full-bleed content like tables)
 * @param {string} variant - 'default' | 'highlight' (optional visual emphasis)
 */
const Card = ({ 
  children, 
  className = '', 
  noPadding = false,
  variant = 'default' 
}) => {
  const baseClass = 'card';
  const variantClass = `card--${variant}`;
  const paddingClass = noPadding ? 'card--no-padding' : '';
  
  return (
    <div className={`${baseClass} ${variantClass} ${paddingClass} ${className}`}>
      {children}
    </div>
  );
};

export default Card;