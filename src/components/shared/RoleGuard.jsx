
import React from 'react';
import { Navigate, useLocation } from 'react-router-dom';
import { useAuth } from '../../hooks/useAuth';

const RoleGuard = ({ children, allowedRoles = [] }) => {
  const { user, loading } = useAuth();
  const location = useLocation();


  if (loading) {
    return <div>Loading...</div>;
  }


  if (!user) {
    return <Navigate to="/login" state={{ from: location }} replace />;
  }


  const normalizeRole = (role) => {
    if (!role) return null;
    return typeof role === 'string' ? role.toLowerCase() : String(role).toLowerCase();
  };

  const rolesArray = Array.isArray(user.roles)
    ? user.roles.map(normalizeRole).filter(Boolean)
    : user.role
    ? [normalizeRole(user.role)]
    : [];


  const normalizedAllowedRoles = allowedRoles.map(role => 
    typeof role === 'string' ? role.toLowerCase() : String(role).toLowerCase()
  );

  if (normalizedAllowedRoles.length > 0 && 
      !rolesArray.some(userRole => normalizedAllowedRoles.includes(userRole))) {
    console.warn(
      `Access denied: user role(s) [${rolesArray.join(', ')}], required: [${normalizedAllowedRoles.join(', ')}]`
    );
    return <Navigate to="/unauthorized" replace />;
  }

  return children;
};

export default RoleGuard;