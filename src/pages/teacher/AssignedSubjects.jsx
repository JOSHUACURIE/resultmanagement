// src/pages/teacher/AssignedSubjects.jsx
import React, { useState, useEffect } from 'react';
import { teacherApi } from '../../api/teacherApi';
import { useAuth } from '../../hooks/useAuth';
import Card from '../../components/ui/Card';
import Button from '../../components/ui/Button';
import Table from '../../components/ui/Table';
import Select from '../../components/ui/Select';
import Tabs from '../../components/ui/Tabs';
import Badge from '../../components/ui/Badge';
import Modal from '../../components/ui/Modal';
import './AssignedSubjects.css';

const AssignedSubjects = () => {
  const { user } = useAuth();
  const [subjects, setSubjects] = useState([]);
  const [assignments, setAssignments] = useState([]);
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [activeTab, setActiveTab] = useState('subjects');
  const [academicYear, setAcademicYear] = useState(new Date().getFullYear().toString());
  const [termId, setTermId] = useState('');
  const [selectedAssignment, setSelectedAssignment] = useState(null);
  const [showAssignmentModal, setShowAssignmentModal] = useState(false);

  // Fetch teacher's assigned subjects
  const fetchSubjects = async () => {
    try {
      setLoading(true);
      const response = await teacherApi.getMySubjects(academicYear, termId);
      if (response.success) {
        setSubjects(response.data || []);
        setError('');
      } else {
        throw new Error(response.message || 'Failed to fetch subjects');
      }
    } catch (err) {
      console.error('Failed to fetch subjects:', err);
      setError('Failed to load your assigned subjects.');
    } finally {
      setLoading(false);
    }
  };

  // Fetch teacher's assignments by term
  const fetchAssignments = async () => {
    try {
      setLoading(true);
      const response = await teacherApi.getMyAssignments(academicYear, termId);
      if (response.success) {
        setAssignments(response.data || []);
        setError('');
      } else {
        throw new Error(response.message || 'Failed to fetch assignments');
      }
    } catch (err) {
      console.error('Failed to fetch assignments:', err);
      setError('Failed to load your assignments.');
    } finally {
      setLoading(false);
    }
  };

  // Fetch teacher's assignment statistics
  const fetchStats = async () => {
    try {
      setLoading(true);
      const response = await teacherApi.getMyAssignmentStats(academicYear, termId);
      if (response.success) {
        setStats(response.data);
        setError('');
      } else {
        throw new Error(response.message || 'Failed to fetch stats');
      }
    } catch (err) {
      console.error('Failed to fetch stats:', err);
      setError('Failed to load statistics.');
    } finally {
      setLoading(false);
    }
  };

  // Load data based on active tab
  useEffect(() => {
    if (user) {
      switch (activeTab) {
        case 'subjects':
          fetchSubjects();
          break;
        case 'assignments':
          fetchAssignments();
          break;
        case 'stats':
          fetchStats();
          break;
        default:
          fetchSubjects();
      }
    }
  }, [user, activeTab, academicYear, termId]);

  const handleSubmitScores = (assignmentId) => {
    window.location.href = `/teacher/submit-scores?assignmentId=${assignmentId}`;
  };

  const handleViewPerformance = (assignmentId) => {
    window.location.href = `/teacher/performance?assignmentId=${assignmentId}`;
  };

  const handleViewStudents = (assignmentId) => {
    window.location.href = `/teacher/students?assignmentId=${assignmentId}`;
  };

  const handleViewAssignmentDetails = async (assignmentId) => {
    try {
      const response = await teacherApi.getMyAssignmentDetails(assignmentId);
      if (response.success) {
        setSelectedAssignment(response.data);
        setShowAssignmentModal(true);
      }
    } catch (err) {
      console.error('Failed to fetch assignment details:', err);
      alert('Failed to load assignment details.');
    }
  };

  // Subjects Tab Columns
  const subjectsColumns = [
    {
      key: 'subject_name',
      header: 'Subject Name',
      render: (value, row) => (
        <div className="subject-info">
          <div className="subject-name">{value}</div>
          <div className="subject-code">{row.subject_code}</div>
          <Badge variant={row.subject_type === 'core' ? 'primary' : 'secondary'}>
            {row.subject_type}
          </Badge>
        </div>
      )
    },
    {
      key: 'assignments',
      header: 'Classes & Streams',
      render: (value, row) => (
        <div className="assignment-list">
          {row.assignments?.map((assignment, index) => (
            <div key={index} className="assignment-item">
              <span className="class-stream">
                {assignment.class.class_name} - {assignment.stream.stream_name}
              </span>
              <span className="term-year">
                {assignment.term.term_name} {assignment.academic_year}
              </span>
            </div>
          ))}
        </div>
      )
    },
    {
      key: 'total_assignments',
      header: 'Total Classes',
      render: (value, row) => `${row.assignments?.length || 0}`
    },
    {
      key: 'actions',
      header: 'Actions',
      render: (_, subject) => (
        <div className="action-buttons">
          <Button
            size="sm"
            variant="primary"
            onClick={() => handleSubmitScores(subject.assignments?.[0]?.assignment_id)}
            disabled={!subject.assignments?.length}
          >
            Submit Scores
          </Button>
          <Button
            size="sm"
            variant="outline"
            onClick={() => handleViewPerformance(subject.assignments?.[0]?.assignment_id)}
            disabled={!subject.assignments?.length}
          >
            Performance
          </Button>
        </div>
      )
    }
  ];

  // Assignments Tab Columns
  const assignmentsColumns = [
    {
      key: 'assignment_subject',
      header: 'Subject',
      render: (value) => (
        <div className="subject-info">
          <div className="subject-name">{value?.subject_name}</div>
          <div className="subject-code">{value?.subject_code}</div>
        </div>
      )
    },
    {
      key: 'assignment_class',
      header: 'Class',
      render: (value, row) => (
        <div className="class-info">
          <div>{value?.class_name}</div>
          <div className="stream-name">{row.assignment_stream?.stream_name}</div>
        </div>
      )
    },
    {
      key: 'assignment_term',
      header: 'Term',
      render: (value) => (
        <div className="term-info">
          <div>{value?.term_name}</div>
          <div className="academic-year">{value?.academic_year}</div>
        </div>
      )
    },
    {
      key: 'actions',
      header: 'Actions',
      render: (_, assignment) => (
        <div className="action-buttons">
          <Button
            size="sm"
            variant="primary"
            onClick={() => handleSubmitScores(assignment.assignment_id)}
          >
            Submit Scores
          </Button>
          <Button
            size="sm"
            variant="outline"
            onClick={() => handleViewStudents(assignment.assignment_id)}
          >
            View Students
          </Button>
          <Button
            size="sm"
            variant="secondary"
            onClick={() => handleViewAssignmentDetails(assignment.assignment_id)}
          >
            Details
          </Button>
        </div>
      )
    }
  ];

  // Stats Tab Content
  const renderStats = () => {
    if (!stats) {
      return <div className="empty-state">Loading statistics...</div>;
    }
    return (
      <div className="stats-grid">
        <Card className="stat-card">
          <h3>Current Year Assignments</h3>
          <div className="stat-number">{stats.current_year_assignments || 0}</div>
        </Card>
        <Card className="stat-card">
          <h3>Total Career Assignments</h3>
          <div className="stat-number">{stats.total_career_assignments || 0}</div>
        </Card>
        {stats.statistics?.map((stat, index) => (
          <Card key={index} className="stat-card">
            <h3>{stat.academic_year} - Term {stat.term_id}</h3>
            <div className="stat-number">{stat.total_assignments}</div>
            <div className="stat-details">
              <div>{stat.total_subjects} subjects</div>
              <div>{stat.total_classes} classes</div>
            </div>
          </Card>
        ))}
      </div>
    );
  };

  return (
    <div className="assigned-subjects">
      <div className="page-header">
        <div className="header-content">
          <h1 className="page-title">My Teaching Assignments</h1>
          <p className="page-subtitle">
            Welcome, <strong>{user?.name}</strong>! Manage your classes and subjects.
          </p>
        </div>
      </div>

      {/* Filters */}
      <Card className="filters-card">
        <div className="filters-row">
          <div className="filter-group">
            <label>Academic Year</label>
            <Select
              value={academicYear}
              onChange={(e) => setAcademicYear(e.target.value)}
            >
              <option value="2024">2024</option>
              <option value="2025">2025</option>
              <option value="2026">2026</option>
            </Select>
          </div>
          <div className="filter-group">
            <label>Term</label>
            <Select
              value={termId}
              onChange={(e) => setTermId(e.target.value)}
            >
              <option value="">All Terms</option>
              <option value="1">Term 1</option>
              <option value="2">Term 2</option>
              <option value="3">Term 3</option>
            </Select>
          </div>
        </div>
      </Card>

      {/* Tabs */}
      <Card className="main-content-card">
        <Tabs
          tabs={[
            { id: 'subjects', label: 'My Subjects' },
            { id: 'assignments', label: 'All Assignments' },
            { id: 'stats', label: 'Statistics' }
          ]}
          activeTab={activeTab}
          onChange={setActiveTab}
        />
        
        <div className="tab-content">
          {error && (
            <div className="error-message">
              <span>{error}</span>
              <Button 
                variant="secondary" 
                size="sm" 
                onClick={() => {
                  switch (activeTab) {
                    case 'subjects': fetchSubjects(); break;
                    case 'assignments': fetchAssignments(); break;
                    case 'stats': fetchStats(); break;
                    default: fetchSubjects();
                  }
                }}
              >
                Retry
              </Button>
            </div>
          )}

          {/* Subjects Tab */}
          {activeTab === 'subjects' && (
            <Table
              columns={subjectsColumns}
              data={subjects}
              loading={loading}
              emptyMessage="You have no assigned subjects for the selected term."
            />
          )}

          {/* Assignments Tab */}
          {activeTab === 'assignments' && (
            <Table
              columns={assignmentsColumns}
              data={assignments}
              loading={loading}
              emptyMessage="No assignments found for the selected term."
            />
          )}

          {/* Stats Tab */}
          {activeTab === 'stats' && renderStats()}
        </div>
      </Card>

      {/* Assignment Details Modal */}
      <Modal
        isOpen={showAssignmentModal}
        onClose={() => setShowAssignmentModal(false)}
        title="Assignment Details"
        size="lg"
      >
        {selectedAssignment && (
          <div className="assignment-details">
            <div className="detail-section">
              <h3>Subject Information</h3>
              <div className="detail-grid">
                <div className="detail-item">
                  <strong>Subject:</strong> {selectedAssignment.assignment_subject?.subject_name}
                </div>
                <div className="detail-item">
                  <strong>Code:</strong> {selectedAssignment.assignment_subject?.subject_code}
                </div>
                <div className="detail-item">
                  <strong>Type:</strong>
                  <Badge variant={selectedAssignment.assignment_subject?.subject_type === 'core' ? 'primary' : 'secondary'}>
                    {selectedAssignment.assignment_subject?.subject_type}
                  </Badge>
                </div>
              </div>
            </div>
            <div className="detail-section">
              <h3>Class Information</h3>
              <div className="detail-grid">
                <div className="detail-item">
                  <strong>Class:</strong> {selectedAssignment.assignment_class?.class_name}
                </div>
                <div className="detail-item">
                  <strong>Stream:</strong> {selectedAssignment.assignment_stream?.stream_name}
                </div>
                <div className="detail-item">
                  <strong>Level:</strong> {selectedAssignment.assignment_class?.class_level}
                </div>
              </div>
            </div>
            <div className="detail-section">
              <h3>Term Information</h3>
              <div className="detail-grid">
                <div className="detail-item">
                  <strong>Term:</strong> {selectedAssignment.assignment_term?.term_name}
                </div>
                <div className="detail-item">
                  <strong>Academic Year:</strong> {selectedAssignment.assignment_term?.academic_year}
                </div>
                <div className="detail-item">
                  <strong>Assigned Date:</strong> {new Date(selectedAssignment.assigned_date).toLocaleDateString()}
                </div>
              </div>
            </div>
            <div className="modal-actions">
              <Button
                variant="primary"
                onClick={() => handleSubmitScores(selectedAssignment.assignment_id)}
              >
                Submit Scores
              </Button>
              <Button
                variant="outline"
                onClick={() => handleViewStudents(selectedAssignment.assignment_id)}
              >
                View Students
              </Button>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
};

export default AssignedSubjects;