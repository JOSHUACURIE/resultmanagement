// src/pages/teacher/SubjectPerformance.jsx
import React, { useState, useEffect } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { get, teacherApi } from '../../api/api';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  Title,
  Tooltip,
  Legend,
  ArcElement
} from 'chart.js';
import { Bar, Line } from 'react-chartjs-2';
import Card from '../../components/ui/Card';
import Button from '../../components/ui/Button';
import './SubjectPerformance.css';

// Register Chart.js components
ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  Title,
  Tooltip,
  Legend,
  ArcElement
);

const SubjectPerformance = () => {
  const navigate = useNavigate();
  const location = useLocation();
  
  const urlParams = new URLSearchParams(location.search);
  const initialSubjectId = urlParams.get('subjectId');

  const [subjects, setSubjects] = useState([]);
  const [selectedSubject, setSelectedSubject] = useState(initialSubjectId || '');
  const [performanceData, setPerformanceData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Fetch assigned subjects
  const fetchSubjects = async () => {
    try {
      // fetch assigned via teacher endpoint
      const id = JSON.parse(localStorage.getItem('user'))?.user?.id || JSON.parse(localStorage.getItem('user'))?.id;
      const res = await teacherApi.getSubjects(id);
      const data = res?.data || res || [];

      const mapped = data.map((a) => ({
        id: a.id,
        name: a.assignment_subject?.subject_name || '—',
        class: a.assignment_class?.class_name ? `${a.assignment_class.class_name} - ${a.assignment_stream?.stream_name || ''}`.trim() : '—',
      }));

      setSubjects(mapped);
      
      if (!initialSubjectId && mapped.length > 0) {
        setSelectedSubject(mapped[0].id);
      }
    } catch (err) {
      setError('Failed to load subjects');
    }
  };

  // Fetch performance data
  const fetchPerformance = async (subjectId) => {
    if (!subjectId) return;
    
    try {
      const data = await get(`/scores/submitted/${subjectId}`);
      setPerformanceData(data);
      setError('');
    } catch (err) {
      setError('Failed to load performance data');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchSubjects();
  }, []);

  useEffect(() => {
    if (selectedSubject) {
      fetchPerformance(selectedSubject);
    } else {
      setPerformanceData(null);
      setLoading(false);
    }
  }, [selectedSubject]);

  const getSubjectName = () => {
    const subject = subjects.find(s => s.id === selectedSubject);
    return subject ? `${subject.name} (${subject.class})` : '';
  };

  // Grade Distribution Chart Data
  const gradeData = {
    labels: ['A', 'B', 'C', 'D', 'E', 'F'],
    datasets: [
      {
        label: 'Number of Students',
        data: performanceData?.gradeDistribution || [0, 0, 0, 0, 0, 0],
        backgroundColor: [
          'rgba(16, 185, 129, 0.7)', // A - success
          'rgba(59, 130, 246, 0.7)', // B - info
          'rgba(245, 158, 11, 0.7)', // C - warning
          'rgba(239, 68, 68, 0.7)',  // D - error
          'rgba(156, 163, 175, 0.7)', // E - neutral
          'rgba(220, 38, 38, 0.7)'   // F - error (darker)
        ],
        borderColor: [
          'rgba(16, 185, 129, 1)',
          'rgba(59, 130, 246, 1)',
          'rgba(245, 158, 11, 1)',
          'rgba(239, 68, 68, 1)',
          'rgba(156, 163, 175, 1)',
          'rgba(220, 38, 38, 1)'
        ],
        borderWidth: 1,
      },
    ],
  };

  const gradeOptions = {
    responsive: true,
    plugins: {
      legend: {
        position: 'top',
      },
      title: {
        display: true,
        text: 'Grade Distribution',
        font: { size: 16 }
      },
    },
    scales: {
      y: {
        beginAtZero: true,
        ticks: {
          stepSize: 1
        }
      }
    }
  };

  // Class Average Trend Data
  const trendData = {
    labels: performanceData?.trend?.map(t => t.term) || [],
    datasets: [
      {
        label: 'Class Average (%)',
        data: performanceData?.trend?.map(t => t.average) || [],
        borderColor: 'rgb(37, 99, 235)',
        backgroundColor: 'rgba(37, 99, 235, 0.1)',
        tension: 0.4,
        fill: true,
      },
    ],
  };

  const trendOptions = {
    responsive: true,
    plugins: {
      legend: {
        position: 'top',
      },
      title: {
        display: true,
        text: 'Class Average Trend',
        font: { size: 16 }
      },
    },
    scales: {
      y: {
        min: 0,
        max: 100,
        ticks: {
          callback: (value) => `${value}%`
        }
      }
    }
  };

  if (loading && !performanceData) {
    return (
      <div className="performance-loading">
        Loading performance data...
      </div>
    );
  }

  return (
    <div className="subject-performance">
      <div className="page-header">
        <h1 className="page-title">Subject Performance</h1>
        <p className="page-subtitle">
          Performance analytics for {getSubjectName() || 'a subject'}
        </p>
      </div>

      {/* Subject Selector */}
      {subjects.length > 1 && (
        <Card className="subject-selector-card">
          <div className="form-field">
            <label htmlFor="subject" className="form-label">
              Select Subject
            </label>
            <select
              id="subject"
              value={selectedSubject}
              onChange={(e) => setSelectedSubject(e.target.value)}
              className="form-select"
            >
              <option value="">-- Choose a subject --</option>
              {subjects.map(subject => (
                <option key={subject.id} value={subject.id}>
                  {subject.name} ({subject.class})
                </option>
              ))}
            </select>
          </div>
        </Card>
      )}

      {error && (
        <Card>
          <div className="performance-error">
            {error}
          </div>
        </Card>
      )}

      {selectedSubject && performanceData && (
        <>
          {/* Summary Stats */}
          <div className="stats-grid">
            <Card variant="highlight">
              <div className="stat-item">
                <span className="stat-label">Class Average</span>
                <p className="stat-value" style={{ color: '#10B981' }}>
                  {performanceData.classAverage}%
                </p>
              </div>
            </Card>

            <Card variant="highlight">
              <div className="stat-item">
                <span className="stat-label">Highest Score</span>
                <p className="stat-value" style={{ color: '#3B82F6' }}>
                  {performanceData.highestScore}%
                </p>
              </div>
            </Card>

            <Card variant="highlight">
              <div className="stat-item">
                <span className="stat-label">Lowest Score</span>
                <p className="stat-value" style={{ color: '#EF4444' }}>
                  {performanceData.lowestScore}%
                </p>
              </div>
            </Card>

            <Card variant="highlight">
              <div className="stat-item">
                <span className="stat-label">Total Students</span>
                <p className="stat-value">
                  {performanceData.totalStudents}
                </p>
              </div>
            </Card>
          </div>

          {/* Charts */}
          <div className="charts-grid">
            <Card>
              <Bar options={gradeOptions} data={gradeData} />
            </Card>

            <Card>
              <Line options={trendOptions} data={trendData} />
            </Card>
          </div>
        </>
      )}

      {!selectedSubject && (
        <Card>
          <div className="no-subject">
            <p>Please select a subject to view performance data.</p>
          </div>
        </Card>
      )}
    </div>
  );
};

export default SubjectPerformance;