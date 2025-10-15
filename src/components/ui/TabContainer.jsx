import React from 'react';
import Tabs from './Tabs';
import TabPanel from './TabPanel';
import './Tabs.css';

const TabsContainer = ({
  tabs,
  activeTab,
  onChange,
  children,
  variant = 'default',
  size = 'md',
  className = '',
  disabled = false
}) => {
  return (
    <div className={`tabs-wrapper ${className}`}>
      <Tabs
        tabs={tabs}
        activeTab={activeTab}
        onChange={onChange}
        variant={variant}
        size={size}
        disabled={disabled}
      />
      <div className="tab-content">
        {React.Children.map(children, (child, index) => {
          if (React.isValidElement(child) && child.type === TabPanel) {
            return React.cloneElement(child, {
              value: activeTab,
              index: tabs[index]?.id || index
            });
          }
          return child;
        })}
      </div>
    </div>
  );
};

export { TabPanel };
export default TabsContainer;