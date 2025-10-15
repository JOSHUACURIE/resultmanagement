import React from 'react';
import { useAuth } from '../../hooks/useAuth';
import { Link, useLocation } from 'react-router-dom';
import {
  FaChalkboardTeacher,
  FaUserGraduate,
  FaBook,
  FaSchool,
  FaChartBar,
  FaComments,
  FaClipboardList,
  FaPaperPlane,
  FaChartLine,
  FaHome,
  FaTachometerAlt,
  FaTimes,
  FaBars
} from 'react-icons/fa';
import './Sidebar.css';

const Sidebar = ({ isOpen, onToggle }) => {
  const { user, loading } = useAuth();
  const location = useLocation();

  const userRole = user?.role ? user.role.toLowerCase() : null;

  const menuConfig = {
    dos: [
      { name: 'Manage Teachers', path: '/dos/manage-teachers', icon: <FaChalkboardTeacher /> },
      { name: 'Manage Students', path: '/dos/manage-students', icon: <FaUserGraduate /> },
      { name: 'Manage Subjects', path: '/dos/manage-subjects', icon: <FaBook /> },
      { name: 'Manage Classes', path: '/dos/manage-classes', icon: <FaSchool /> },
      { name: 'Analyze Results', path: '/dos/analyze-results', icon: <FaChartBar /> },
      { name: 'Complaints & Suggestions', path: '/dos/complaints', icon: <FaComments /> },
    ],
    teacher: [
      { name: 'Assigned Subjects', path: '/teacher/assigned-subjects', icon: <FaClipboardList /> },
      { name: 'Submit Scores', path: '/teacher/submit-scores', icon: <FaBook /> },
      { name: 'Message Admin', path: '/teacher/message-admin', icon: <FaPaperPlane /> },
      { name: 'Subject Performance', path: '/teacher/performance', icon: <FaChartLine /> },
    ],
    principal: [
      { name: 'View Students', path: '/principal/students', icon: <FaUserGraduate /> },
      { name: 'View Teachers', path: '/principal/teachers', icon: <FaChalkboardTeacher /> },
      { name: 'View Performance', path: '/principal/performance', icon: <FaChartBar /> },
      { name: 'View Complaints', path: '/principal/complaints', icon: <FaComments /> },
    ],
  };

  const menuItems = userRole ? menuConfig[userRole] || [] : [];

  const roleConfig = {
    dos: { displayName: 'DOS Dashboard', icon: <FaTachometerAlt /> },
    teacher: { displayName: 'Teacher Dashboard', icon: <FaChalkboardTeacher /> },
    principal: { displayName: 'Principal Dashboard', icon: <FaHome /> },
  };

  const currentRole = userRole ? roleConfig[userRole] : null;

  const handleLinkClick = () => {
    if (window.innerWidth <= 768) {
      onToggle(false);
    }
  };

  if (loading) {
    return (
      <aside className={`sidebar ${isOpen ? 'sidebar-open' : ''}`}>
        <div className="sidebar-header">
          <div className="sidebar-logo">
            <FaSchool className="logo-icon" />
            <h2>School RMS</h2>
          </div>
          <div className="sidebar-loading">
            <div className="loading-spinner"></div>
            <p>Loading...</p>
          </div>
        </div>
      </aside>
    );
  }

  return (
    <>
      {/* 🔹 Hamburger button for mobile */}
      <button className="sidebar-toggle-btn" onClick={() => onToggle(!isOpen)}>
        {isOpen ? <FaTimes /> : <FaBars />}
      </button>

      {/* 🔹 Dim overlay when sidebar is open on mobile */}
      {isOpen && <div className="sidebar-overlay" onClick={() => onToggle(false)} />}

      {/* 🔹 Sidebar content */}
      <aside className={`sidebar ${isOpen ? 'sidebar-open' : ''}`}>
        <div className="sidebar-header">
          <div className="sidebar-logo">
            <FaSchool className="logo-icon" />
            <h2>School RMS</h2>
          </div>

          {currentRole && (
            <div className="role-info">
              <span className="role-icon">{currentRole.icon}</span>
              <p>{currentRole.displayName}</p>
            </div>
          )}
        </div>

        <nav className="sidebar-nav">
          <ul>
            {menuItems.map((item) => (
              <li key={item.path}>
                <Link
                  to={item.path}
                  className={`nav-link ${location.pathname === item.path ? 'active' : ''}`}
                  onClick={handleLinkClick}
                >
                  <span className="nav-icon">{item.icon}</span>
                  <span className="nav-text">{item.name}</span>
                </Link>
              </li>
            ))}
          </ul>
        </nav>

        {user && (
          <div className="sidebar-footer">
            <div className="user-info">
              <div className="user-avatar">
                <FaUserGraduate />
              </div>
              <div className="user-details">
                <span className="user-name">{user.name || 'User'}</span>
                <span className="user-role">{user.role}</span>
              </div>
            </div>
          </div>
        )}
      </aside>
    </>
  );
};

export default Sidebar;
