// src/pages/teacher/Dashboard.jsx
import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { get } from '../../api/api';
import { useAuth } from '../../hooks/useAuth';
import Card from '../../components/ui/Card';
import Badge from '../../components/ui/Badge';
import Button from '../../components/ui/Button';
import Table from '../../components/ui/Table';
import './Dashboard.css';

const TeacherDashboard = () => {
  const { user } = useAuth();
  const navigate = useNavigate();
  
  // State
  const [stats, setStats] = useState({
    assignedSubjects: 0,
    pendingSubmissions: 0,
    totalStudents: 0
  });
  
  const [recentScores, setRecentScores] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Fetch dashboard data
  useEffect(() => {
    const fetchData = async () => {
      try {
        const [
          subjectsRes,
          pendingRes,
          studentsRes
        ] = await Promise.all([
          teacherApi.getSubjects(user?.id || user?.user_id),
          get('/results/pending-submissions'),
          get('/students')
        ]);

        setStats({
          assignedSubjects: subjectsRes.length || 0,
          pendingSubmissions: pendingRes.length || 0,
          totalStudents: studentsRes.length || 0
        });

        // Prepare recent scores for table
        setRecentScores(pendingRes.slice(0, 5).map(item => ({
          id: item.id,
          subject: item.subjectName,
          class: item.className,
          dueDate: new Date(item.dueDate).toLocaleDateString(),
          status: 'Pending'
        })));

      } catch (err) {
        console.error('Failed to load dashboard ', err);
        setError('Failed to load dashboard data. Please try again later.');
      } finally {
        setLoading(false);
      }
    };

    if (user) fetchData();
  }, [user]);

  // Quick actions
  const handleViewSubjects = () => navigate('/teacher/assigned-subjects');
  const handleSubmitScores = () => navigate('/teacher/submit-scores');
  const handleMessageAdmin = () => navigate('/teacher/message-admin');
  const handleViewPerformance = () => navigate('/teacher/performance');

  // Table columns for pending submissions
  const pendingColumns = [
    { key: 'subject', header: 'Subject' },
    { key: 'class', header: 'Class' },
    { key: 'dueDate', header: 'Due Date' },
    { 
      key: 'status', 
      header: 'Status',
      render: () => <Badge variant="warning">Pending</Badge>
    }
  ];

  if (loading) {
    return (
      <div className="dashboard-loading">
        Loading your dashboard...
      </div>
    );
  }

  if (error) {
    return (
      <Card>
        <div className="dashboard-error">
          <p>{error}</p>
          <Button 
            variant="primary" 
            size="sm" 
            onClick={() => window.location.reload()}
            className="mt-4"
          >
            Retry
          </Button>
        </div>
      </Card>
    );
  }

  return (
    <div className="teacher-dashboard">
      {/* Header */}
      <div className="dashboard-header">
        <h1 className="dashboard-title">Teacher Dashboard</h1>
        <p className="dashboard-subtitle">
          Welcome back, <strong>{user?.name}</strong>! Here's your teaching overview.
        </p>
      </div>

      {/* Stats Cards */}
      <div className="stats-grid">
        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="teacher" size="sm">Subjects</Badge>
            <h3 className="stat-label">Assigned Subjects</h3>
            <p className="stat-value">{stats.assignedSubjects}</p>
          </div>
        </Card>

        {/* In DosDashboard.jsx */}
<Card variant="highlight">
  <div className="stat-item dos-highlight"> {/* ← Added class */}
    <Badge variant="dos" size="sm">DOS</Badge>
    <h3 className="stat-label">Pending Scores</h3>
    <p className="stat-value">{stats.pendingScores}</p>
  </div>
</Card>

        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="info" size="sm">Students</Badge>
            <h3 className="stat-label">Total Students</h3>
            <p className="stat-value">{stats.totalStudents}</p>
          </div>
        </Card>
      </div>

      {/* Quick Actions */}
      <div className="quick-actions-section">
        <h2 className="section-title">Quick Actions</h2>
        <div className="quick-actions-grid">
          <Button 
            onClick={handleViewSubjects}
            className="quick-action-btn"
          >
            View Assigned Subjects
          </Button>
          <Button 
            onClick={handleSubmitScores}
            className="quick-action-btn"
          >
            Submit Scores
          </Button>
          <Button 
            onClick={handleViewPerformance}
            className="quick-action-btn"
          >
            View Subject Performance
          </Button>
          <Button 
            onClick={handleMessageAdmin}
            className="quick-action-btn"
          >
            Message Admin (DOS)
          </Button>
        </div>
      </div>

      {/* Pending Submissions */}
      {stats.pendingSubmissions > 0 && (
        <div className="pending-section">
          <Card>
            <div className="pending-header">
              <h3 className="pending-title">Pending Score Submissions</h3>
              <Button 
                variant="outline" 
                size="sm" 
                onClick={handleSubmitScores}
              >
                Submit Now
              </Button>
            </div>
            <Table 
              columns={pendingColumns} 
              data={recentScores} 
              emptyMessage="No pending submissions"
            />
          </Card>
        </div>
      )}
    </div>
  );
};

export default TeacherDashboard;