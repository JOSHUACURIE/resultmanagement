// src/components/layout/DashboardLayout.jsx
import { useState } from 'react';
import { Outlet } from 'react-router-dom';
import Navbar from './Navbar';
import Sidebar from './SideBar';
import './DashboardLayout.css';

const DashboardLayout = () => {
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

  const handleToggleSidebar = (state) => {
    // If explicit boolean provided (from child), use it; otherwise toggle
    setIsSidebarOpen(state !== undefined ? state : !isSidebarOpen);
  };

  return (
    <div className={`dashboard-layout ${isSidebarOpen ? 'sidebar-active' : ''}`}>
      {/* Sidebar */}
      <Sidebar isOpen={isSidebarOpen} onToggle={handleToggleSidebar} />

      {/* Main Section */}
      <div className="dashboard-main">
        <Navbar />
        <main className="dashboard-content">
          <Outlet />
        </main>
      </div>
    </div>
  );
};

export default DashboardLayout;
