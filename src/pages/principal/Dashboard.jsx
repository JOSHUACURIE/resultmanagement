
import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { get } from '../../api/api';
import Card from '../../components/ui/Card';
import Badge from '../../components/ui/Badge';
import Button from '../../components/ui/Button';
import Table from '../../components/ui/Table';
import './Dashboard.css';

const PrincipalDashboard = () => {
  const navigate = useNavigate();
  
  // State
  const [stats, setStats] = useState({
    totalStudents: 0,
    totalTeachers: 0,
    totalSubjects: 0,
    passRate: 0
  });
  
  const [topSubjects, setTopSubjects] = useState([]);
  const [recentComplaints, setRecentComplaints] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Fetch dashboard data
  useEffect(() => {
    const fetchData = async () => {
      try {
        const [
          studentsRes,
          teachersRes,
          subjectsRes,
          performanceRes,
          complaintsRes
        ] = await Promise.all([
          get('/students'),
          get('/teachers'),
          get('/subjects'),
          get('/results/school-performance'),
          get('/complaints/recent')
        ]);

        setStats({
          totalStudents: studentsRes.length || 0,
          totalTeachers: teachersRes.length || 0,
          totalSubjects: subjectsRes.length || 0,
          passRate: performanceRes.passRate || 0
        });

        setTopSubjects(performanceRes.topSubjects || []);
        setRecentComplaints(complaintsRes.slice(0, 5) || []);

      } catch (err) {
        console.error('Failed to load dashboard ', err);
        setError('Failed to load dashboard data. Please try again later.');
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  // Quick actions
  const handleViewStudents = () => navigate('/principal/students');
  const handleViewTeachers = () => navigate('/principal/teachers');
  const handleViewPerformance = () => navigate('/principal/performance');
  const handleViewComplaints = () => navigate('/principal/complaints');

  // Table columns for complaints
  const complaintColumns = [
    { 
      key: 'fromName', 
      header: 'From',
      render: (name, row) => (
        <div>
          <div>{name}</div>
          <div className="complaint-role">{row.fromRole}</div>
        </div>
      )
    },
    { 
      key: 'category', 
      header: 'Category',
      render: (category) => (
        <Badge variant={getCategoryVariant(category)}>
          {category}
        </Badge>
      )
    },
    { 
      key: 'message', 
      header: 'Message',
      render: (msg) => (
        <span className="complaint-message">
          {msg.length > 40 ? `${msg.substring(0, 40)}...` : msg}
        </span>
      )
    },
    { 
      key: 'submittedAt', 
      header: 'Date',
      render: (date) => new Date(date).toLocaleDateString()
    }
  ];

  // Get category badge variant
  const getCategoryVariant = (category) => {
    const lower = category.toLowerCase();
    if (lower.includes('academic')) return 'info';
    if (lower.includes('behavior') || lower.includes('discipline')) return 'warning';
    if (lower.includes('facility') || lower.includes('infrastructure')) return 'neutral';
    return 'primary';
  };

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
    <div className="principal-dashboard">
      {/* Header */}
      <div className="dashboard-header">
        <h1 className="dashboard-title">Principal Dashboard</h1>
        <p className="dashboard-subtitle">
          Overview of school performance and key metrics
        </p>
      </div>

      {/* Stats Cards */}
      <div className="stats-grid">
        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="info" size="sm">Students</Badge>
            <h3 className="stat-label">Total Students</h3>
            <p className="stat-value">{stats.totalStudents}</p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="teacher" size="sm">Teachers</Badge>
            <h3 className="stat-label">Total Teachers</h3>
            <p className="stat-value">{stats.totalTeachers}</p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="primary" size="sm">Subjects</Badge>
            <h3 className="stat-label">Total Subjects</h3>
            <p className="stat-value">{stats.totalSubjects}</p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="success" size="sm">Pass Rate</Badge>
            <h3 className="stat-label">School Pass Rate</h3>
            <p className="stat-value">{stats.passRate}%</p>
          </div>
        </Card>
      </div>

      {/* Quick Actions */}
      <div className="quick-actions-section">
        <h2 className="section-title">Quick Actions</h2>
        <div className="quick-actions-grid">
          <Button 
            onClick={handleViewStudents}
            className="quick-action-btn"
          >
            View Students
          </Button>
          <Button 
            onClick={handleViewTeachers}
            className="quick-action-btn"
          >
            View Teachers
          </Button>
          <Button 
            onClick={handleViewPerformance}
            className="quick-action-btn"
          >
            View Performance
          </Button>
          <Button 
            onClick={handleViewComplaints}
            className="quick-action-btn"
          >
            View Complaints
          </Button>
        </div>
      </div>

      {/* Top Performing Subjects */}
      {topSubjects.length > 0 && (
        <div className="top-subjects-section">
          <Card>
            <h3 className="section-title">Top Performing Subjects</h3>
            <div className="subjects-grid">
              {topSubjects.map((subject, index) => (
                <div key={subject.id} className="subject-card">
                  <div className="subject-rank">#{index + 1}</div>
                  <div className="subject-info">
                    <h4 className="subject-name">{subject.name}</h4>
                    <p className="subject-class">{subject.class}</p>
                  </div>
                  <div className="subject-average">
                    <span className="average-value">{subject.average}%</span>
                  </div>
                </div>
              ))}
            </div>
          </Card>
        </div>
      )}

      {/* Recent Complaints */}
      {recentComplaints.length > 0 && (
        <div className="recent-complaints-section">
          <Card>
            <div className="complaints-header">
              <h3 className="section-title">Recent Complaints</h3>
              <Button 
                variant="outline" 
                size="sm" 
                onClick={handleViewComplaints}
              >
                View All
              </Button>
            </div>
            <Table 
              columns={complaintColumns} 
              data={recentComplaints} 
              emptyMessage="No recent complaints"
            />
          </Card>
        </div>
      )}
    </div>
  );
};

export default PrincipalDashboard;