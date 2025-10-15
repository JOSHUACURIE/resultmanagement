import React from 'react';
import './Tabs.css';

const TabPanel = ({ 
  children, 
  value, 
  index, 
  className = '',
  lazy = false,
  keepMounted = false 
}) => {
  const isActive = value === index;
  
  // For lazy loading: only render content when active
  if (lazy && !isActive) {
    return null;
  }

  // For keepMounted: render but hide
  const shouldRender = isActive || keepMounted;
  
  if (!shouldRender) {
    return null;
  }

  return (
    <div 
      className={`tab-panel ${isActive ? 'tab-panel-active' : 'tab-panel-hidden'} ${className}`}
      role="tabpanel"
      hidden={!isActive}
    >
      {children}
    </div>
  );
};

export default TabPanel;