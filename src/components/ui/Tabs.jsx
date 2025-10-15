import React from 'react';
import './Tabs.css';

const Tabs = ({ 
  tabs, 
  activeTab, 
  onChange, 
  variant = 'default',
  size = 'md',
  className = '',
  disabled = false 
}) => {
  const handleTabClick = (tabId) => {
    if (!disabled && onChange) {
      onChange(tabId);
    }
  };

  return (
    <div className={`tabs-container tabs-${variant} ${className}`}>
      <div className={`tabs-header tabs-size-${size}`}>
        {tabs.map((tab) => (
          <button
            key={tab.id}
            className={`tab-button ${
              activeTab === tab.id ? 'tab-active' : ''
            } ${tab.disabled ? 'tab-disabled' : ''}`}
            onClick={() => !tab.disabled && handleTabClick(tab.id)}
            disabled={disabled || tab.disabled}
            type="button"
          >
            {tab.icon && <span className="tab-icon">{tab.icon}</span>}
            {tab.label}
            {tab.badge !== undefined && (
              <span className="tab-badge">{tab.badge}</span>
            )}
          </button>
        ))}
      </div>
    </div>
  );
};

export default Tabs;