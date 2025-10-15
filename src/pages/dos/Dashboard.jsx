import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { get } from '../../api/api';
import Card from '../../components/ui/Card';
import Badge from '../../components/ui/Badge';
import Button from '../../components/ui/Button';
import Table from '../../components/ui/Table';
import './Dashboard.css';

const DosDashboard = () => {
  const navigate = useNavigate();
  
  const [stats, setStats] = useState({
    totalTeachers: 0,
    totalStudents: 0,
    pendingScores: 0
  });
  
  const [recentScores, setRecentScores] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Fetch dashboard data
  useEffect(() => {
    const fetchData = async () => {
      try {
        // Get current term and academic year for results filtering
        const currentYear = new Date().getFullYear().toString();
        
        let currentTermId = '';
        let defaultClassId = '';

        try {
          // Fetch terms - handle different response structures
          const termsRes = await get('/terms');
          // Handle different API response structures
          const termsData = termsRes.data || termsRes.terms || Array.isArray(termsRes) ? termsRes : [];
          const termsArray = Array.isArray(termsData) ? termsData : [];
          
          // Find current term - adjust property names based on your API
          const currentTerm = termsArray.find(term => 
            term.is_active || term.status === 'active' || term.is_current
          ) || termsArray[0];
          
          currentTermId = currentTerm?.term_id || '';

        } catch (termsError) {
          console.warn('Could not fetch terms:', termsError);
          // Continue without term data
        }

        try {
          // Fetch classes - handle different response structures
          const classesRes = await get('/classes');
          const classesData = classesRes.data || classesRes.classes || Array.isArray(classesRes) ? classesRes : [];
          const classesArray = Array.isArray(classesData) ? classesData : [];
          defaultClassId = classesArray[0]?.class_id || '';
        } catch (classesError) {
          console.warn('Could not fetch classes:', classesError);
          // Continue without class data
        }

        // Fetch all data in parallel with error handling for each request
        const requests = [
          get('/teachers').catch(err => {
            console.warn('Failed to fetch teachers:', err);
            return [];
          }),
          get('/students').catch(err => {
            console.warn('Failed to fetch students:', err);
            return [];
          })
        ];

        // Only add results request if we have the required parameters
        if (currentTermId && defaultClassId) {
          requests.push(
            get(`/results/submitted?term_id=${currentTermId}&class_id=${defaultClassId}&academic_year=${currentYear}`)
              .catch(err => {
                console.warn('Failed to fetch submitted results:', err);
                return { success: false, data: [] };
              })
          );
        } else {
          requests.push(Promise.resolve({ success: false, data: [] }));
        }

        const [
          teachersRes,
          studentsRes,
          submittedResultsRes
        ] = await Promise.all(requests);

        // Extract data from different response structures
        const teachersData = teachersRes.data || Array.isArray(teachersRes) ? teachersRes : [];
        const studentsData = studentsRes.data || Array.isArray(studentsRes) ? studentsRes : [];
        
        const teachersArray = Array.isArray(teachersData) ? teachersData : [];
        const studentsArray = Array.isArray(studentsData) ? studentsData : [];

        // Calculate pending scores
        const totalStudents = studentsArray.length;
        let studentsWithScores = 0;
        let submittedResultsData = [];

        if (submittedResultsRes && submittedResultsRes.success) {
          submittedResultsData = submittedResultsRes.data || [];
          studentsWithScores = Array.isArray(submittedResultsData) ? submittedResultsData.length : 0;
        }

        const pendingScoresCount = Math.max(0, totalStudents - studentsWithScores);

        setStats({
          totalTeachers: teachersArray.length,
          totalStudents: totalStudents,
          pendingScores: pendingScoresCount
        });

        // Format recent submitted scores
        if (Array.isArray(submittedResultsData) && submittedResultsData.length > 0) {
          const recentSubmittedScores = submittedResultsData.slice(0, 5).map(result => {
            const subjectScores = result.subject_scores || [];
            const latestSubject = subjectScores.length > 0 
              ? subjectScores[subjectScores.length - 1] 
              : { subject_name: 'No subjects', score: 'N/A' };
            
            return {
              id: result.student_id || result.admission_number,
              student: result.fullname || 'Unknown Student',
              subject: latestSubject.subject_name || 'Unknown Subject',
              score: latestSubject.score || 'N/A',
              submittedAt: new Date().toLocaleDateString()
            };
          });
          setRecentScores(recentSubmittedScores);
        } else {
          setRecentScores([]);
        }

      } catch (err) {
        console.error('Failed to load dashboard data:', err);
        setError('Failed to load dashboard data. Please try again later.');
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  // Navigation handlers
  const handleManageTeachers = () => navigate('/dos/manage-teachers');
  const handleManageStudents = () => navigate('/dos/manage-students');
  const handleAnalyzeResults = () => navigate('/dos/analyze-results');
  const handleViewSubmittedResults = () => navigate('/dos/submitted-results');

  // Table columns
  const scoreColumns = [
    { key: 'student', header: 'Student' },
    { key: 'subject', header: 'Subject' },
    { 
      key: 'score', 
      header: 'Score',
      render: (score) => score === 'N/A' ? 
        <Badge variant="warning" size="sm">Pending</Badge> : 
        <span className="font-semibold">{score}</span>
    },
    { key: 'submittedAt', header: 'Submitted' }
  ];

  if (loading) {
    return (
      <div className="dashboard-loading">
        <div className="loading-spinner"></div>
        Loading dashboard...
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
    <div className="dos-dashboard">
      {/* Header */}
      <div className="dashboard-header">
        <h1 className="dashboard-title">Director of Studies Dashboard</h1>
        <p className="dashboard-subtitle">
          Overview of school academic performance
        </p>
      </div>

      {/* Stats Grid */}
      <div className="stats-grid">
        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="dos" size="sm">DOS</Badge>
            <h3 className="stat-label">Total Teachers</h3>
            <p className="stat-value">{stats.totalTeachers}</p>
            <p className="stat-change">Active staff members</p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="info" size="sm">Students</Badge>
            <h3 className="stat-label">Total Students</h3>
            <p className="stat-value">{stats.totalStudents}</p>
            <p className="stat-change">Currently enrolled</p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <Badge variant="warning" size="sm">Pending</Badge>
            <h3 className="stat-label">Pending Scores</h3>
            <p className="stat-value">{stats.pendingScores}</p>
            <p className="stat-change">Awaiting submission</p>
          </div>
        </Card>
      </div>

      {/* Quick Actions */}
      <div className="quick-actions-section">
        <h2 className="section-title">Quick Actions</h2>
        <div className="quick-actions-grid">
          <Button 
            onClick={handleManageTeachers}
            className="quick-action-btn"
          >
            Manage Teachers
          </Button>
          <Button 
            onClick={handleManageStudents}
            className="quick-action-btn"
          >
            Manage Students
          </Button>
          <Button 
            onClick={handleAnalyzeResults}
            className="quick-action-btn"
          >
            Analyze Results
          </Button>
          <Button 
            onClick={handleViewSubmittedResults}
            className="quick-action-btn"
          >
            View Submitted Results
          </Button>
        </div>
      </div>

      {/* Recent Activity */}
      <div className="recent-activity-section">
        <div className="activity-grid">
          {/* Recent Submitted Scores */}
          <Card>
            <div className="activity-header">
              <h3 className="activity-title">
                Recently Submitted Scores
                {recentScores.length > 0 && (
                  <Badge variant="info" size="sm" className="ml-2">
                    {recentScores.length}
                  </Badge>
                )}
              </h3>
              <Button 
                variant="outline" 
                size="sm" 
                onClick={handleViewSubmittedResults}
              >
                View All
              </Button>
            </div>
            <Table 
              columns={scoreColumns} 
              data={recentScores} 
              emptyMessage="No recently submitted scores found"
            />
          </Card>
        </div>
      </div>
    </div>
  );
};

export default DosDashboard;