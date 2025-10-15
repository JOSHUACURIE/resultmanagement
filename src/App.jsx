// src/App.jsx
import React from 'react';
import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { AuthProvider } from './context/AuthContext';
import RoleGuard from './components/shared/RoleGuard';
import { ROLES } from './utils/roles';

// Auth Pages
import Login from './pages/auth/Login';

// DOS Pages
import DosDashboard from './pages/dos/Dashboard';
import ManageTeachers from './pages/dos/ManageTeachers';
import ManageStudents from './pages/dos/ManageStudents';
import ManageSubjects from './pages/dos/ManageSubjects';
import AnalyzeResults from './pages/dos/AnalyzeResults';
import ManageClasses from './pages/dos/ManageClasses';
import ComplaintsSuggestions from './pages/dos/ComplaintsSuggestions';
import SubmittedResultsView from './pages/dos/SubmittedResultsView';
// Teacher Pages
import TeacherDashboard from './pages/teacher/Dashboard';
import AssignedSubjects from './pages/teacher/AssignedSubjects';
import SubmitScores from './pages/teacher/SubmitScores';
import MessageAdmin from './pages/teacher/MessageAdmin';
import SubjectPerformance from './pages/teacher/SubjectPerformance';

// Principal Pages
import PrincipalDashboard from './pages/principal/Dashboard';
import ViewStudents from './pages/principal/ViewStudents';
import ViewTeachers from './pages/principal/ViewTeachers';
import ViewPerformance from './pages/principal/ViewPerformance';
import ViewComplaints from './pages/principal/ViewComplaints';

// Layout
import DashboardLayout from './components/layout/DashboardLayout';

function App() {
  return (
    <BrowserRouter>
      <AuthProvider>
        <Routes>
          {/* =======================
              PUBLIC ROUTES
          ======================= */}
          <Route path="/login" element={<Login />} />

          {/* =======================
              DOS ROUTES
          ======================= */}
          <Route
            path="/dos/*"
            element={
              <RoleGuard allowedRoles={[ROLES.dos]}>
                <DashboardLayout />
              </RoleGuard>
            }
          >
            {/* Redirect /dos → /dos/dashboard */}
            <Route index element={<Navigate to="dashboard" replace />} />
            <Route path="dashboard" element={<DosDashboard />} />
            <Route path="manage-teachers" element={<ManageTeachers />} />
            <Route path="manage-students" element={<ManageStudents />} />
            <Route path="manage-classes" element={<ManageClasses />} />
            <Route path="manage-subjects" element={<ManageSubjects />} />
            <Route path="analyze-results" element={<AnalyzeResults />} />
            <Route path="complaints" element={<ComplaintsSuggestions />} />
            <Route path="submitted-results" element={<SubmittedResultsView />} />
          </Route>

          {/* =======================
              TEACHER ROUTES
          ======================= */}
          <Route
            path="/teacher/*"
            element={
              <RoleGuard allowedRoles={[ROLES.teacher]}>
                <DashboardLayout />
              </RoleGuard>
            }
          >
            {/* Redirect /teacher → /teacher/dashboard */}
            <Route index element={<Navigate to="dashboard" replace />} />
            <Route path="dashboard" element={<TeacherDashboard />} />
            <Route path="assigned-subjects" element={<AssignedSubjects />} />
            <Route path="submit-scores" element={<SubmitScores />} />
            <Route path="message-admin" element={<MessageAdmin />} />
            <Route path="performance" element={<SubjectPerformance />} />
          </Route>

          {/* =======================
              PRINCIPAL ROUTES
          ======================= */}
          <Route
            path="/principal/*"
            element={
              <RoleGuard allowedRoles={[ROLES.principal]}>
                <DashboardLayout />
              </RoleGuard>
            }
          >
            {/* Redirect /principal → /principal/dashboard */}
            <Route index element={<Navigate to="dashboard" replace />} />
            <Route path="dashboard" element={<PrincipalDashboard />} />
            <Route path="students" element={<ViewStudents />} />
            <Route path="teachers" element={<ViewTeachers />} />
            <Route path="performance" element={<ViewPerformance />} />
            <Route path="complaints" element={<ViewComplaints />} />
          </Route>

          {/* =======================
              DEFAULT & FALLBACK
          ======================= */}
          <Route path="/" element={<Navigate to="/login" replace />} />
          <Route path="*" element={<Navigate to="/login" replace />} />
        </Routes>
      </AuthProvider>
    </BrowserRouter>
  );
}

export default App;
