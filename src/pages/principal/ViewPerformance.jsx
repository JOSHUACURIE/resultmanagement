// src/pages/principal/ViewPerformance.jsx
import React, { useState, useEffect } from 'react';
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
import { Bar, Line} from 'react-chartjs-2';
import Card from '../../components/ui/Card';
import './ViewPerformance.css';

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

const ViewPerformance = () => {
  const [performanceData, setPerformanceData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Mock API call - replace with your actual API
  const fetchPerformanceData = async () => {
    try {
      // Replace this with your actual API call:
      // const data = await get('/results/school-performance');
      
      // Mock data for demonstration
      const mockData = {
        schoolPassRate: 87.5,
        gradeDistribution: {
          labels: ['A', 'B', 'C', 'D', 'E', 'F'],
          counts: [120, 180, 95, 45, 20, 10]
        },
        performanceByClass: {
          labels: ['SS3', 'SS2', 'SS1', 'JS3', 'JS2', 'JS1'],
          averages: [82.3, 79.8, 76.5, 74.2, 71.8, 68.9]
        },
        topSubjects: [
          { name: 'Mathematics', average: 85.2 },
          { name: 'English', average: 82.7 },
          { name: 'Physics', average: 80.1 }
        ],
        bottomSubjects: [
          { name: 'Literature', average: 62.3 },
          { name: 'History', average: 65.8 },
          { name: 'Geography', average: 68.4 }
        ],
        termTrend: {
          terms: ['Term 1', 'Term 2', 'Term 3'],
          averages: [72.5, 78.3, 81.2]
        }
      };
      
      setPerformanceData(mockData);
      setError('');
    } catch (err) {
      setError('Failed to load performance data. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchPerformanceData();
  }, []);

  if (loading && !performanceData) {
    return (
      <div className="performance-loading">
        Loading school performance data...
      </div>
    );
  }

  if (error) {
    return (
      <Card>
        <div className="performance-error">
          {error}
        </div>
      </Card>
    );
  }

  // Grade Distribution Chart
  const gradeData = {
    labels: performanceData?.gradeDistribution.labels || [],
    datasets: [
      {
        label: 'Number of Students',
        data: performanceData?.gradeDistribution.counts || [],
        backgroundColor: [
          'rgba(16, 185, 129, 0.7)', // A
          'rgba(59, 130, 246, 0.7)', // B
          'rgba(245, 158, 11, 0.7)', // C
          'rgba(239, 68, 68, 0.7)',  // D
          'rgba(156, 163, 175, 0.7)', // E
          'rgba(220, 38, 38, 0.7)'   // F
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
      legend: { position: 'top' },
      title: { display: true, text: 'Grade Distribution', font: { size: 16 } }
    },
    scales: {
      y: { beginAtZero: true, ticks: { stepSize: 20 } }
    }
  };

  // Performance by Class Chart
  const classData = {
    labels: performanceData?.performanceByClass.labels || [],
    datasets: [
      {
        label: 'Class Average (%)',
        data: performanceData?.performanceByClass.averages || [],
        backgroundColor: 'rgba(37, 99, 235, 0.7)',
        borderColor: 'rgba(37, 99, 235, 1)',
        borderWidth: 1,
      },
    ],
  };

  const classOptions = {
    indexAxis: 'y',
    responsive: true,
    plugins: {
      legend: { position: 'top' },
      title: { display: true, text: 'Performance by Class', font: { size: 16 } }
    },
    scales: {
      x: { min: 0, max: 100, ticks: { callback: (value) => `${value}%` } }
    }
  };

  // Term Trend Chart
  const trendData = {
    labels: performanceData?.termTrend.terms || [],
    datasets: [
      {
        label: 'School Average (%)',
        data: performanceData?.termTrend.averages || [],
        borderColor: 'rgb(16, 185, 129)',
        backgroundColor: 'rgba(16, 185, 129, 0.1)',
        tension: 0.4,
        fill: true,
        pointRadius: 5,
        pointBackgroundColor: 'rgb(16, 185, 129)'
      },
    ],
  };

  const trendOptions = {
    responsive: true,
    plugins: {
      legend: { position: 'top' },
      title: { display: true, text: 'Term-over-Term Performance Trend', font: { size: 16 } }
    },
    scales: {
      y: { min: 0, max: 100, ticks: { callback: (value) => `${value}%` } }
    }
  };

  return (
    <div className="view-performance">
      <div className="page-header">
        <h1 className="page-title">School Performance</h1>
        <p className="page-subtitle">
          School-wide pass rate: <strong>{performanceData?.schoolPassRate || 0}%</strong>
        </p>
      </div>

      {/* Summary Stats */}
      <div className="stats-grid">
        <Card variant="highlight">
          <div className="stat-item">
            <span className="stat-label">School Pass Rate</span>
            <p className="stat-value" style={{ color: '#10B981' }}>
              {performanceData?.schoolPassRate || 0}%
            </p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <span className="stat-label">Total Students</span>
            <p className="stat-value">
              {performanceData?.gradeDistribution.counts.reduce((a, b) => a + b, 0) || 0}
            </p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <span className="stat-label">Top Subject Avg</span>
            <p className="stat-value" style={{ color: '#3B82F6' }}>
              {performanceData?.topSubjects[0]?.average || 0}%
            </p>
          </div>
        </Card>

        <Card variant="highlight">
          <div className="stat-item">
            <span className="stat-label">Improvement</span>
            <p className="stat-value" style={{ color: '#F59E0B' }}>
              +{((performanceData?.termTrend.averages[2] || 0) - (performanceData?.termTrend.averages[0] || 0)).toFixed(1)}%
            </p>
          </div>
        </Card>
      </div>

      {/* Charts Grid */}
      <div className="charts-grid">
        <Card>
          <Bar options={gradeOptions} data={gradeData} />
        </Card>

        <Card>
          <Bar options={classOptions} data={classData} />
        </Card>
      </div>

      {/* Trend Chart */}
      <div className="trend-chart">
        <Card>
          <Line options={trendOptions} data={trendData} />
        </Card>
      </div>

      {/* Top & Bottom Subjects */}
      <div className="subjects-grid">
        <Card>
          <h3 className="section-title">Top Performing Subjects</h3>
          <div className="subjects-list">
            {performanceData?.topSubjects.map((subject, index) => (
              <div key={subject.name} className="subject-item">
                <span className="subject-rank">#{index + 1}</span>
                <span className="subject-name">{subject.name}</span>
                <span className="subject-average" style={{ color: '#10B981' }}>
                  {subject.average}%
                </span>
              </div>
            ))}
          </div>
        </Card>

        <Card>
          <h3 className="section-title">Subjects Needing Attention</h3>
          <div className="subjects-list">
            {performanceData?.bottomSubjects.map((subject, index) => (
              <div key={subject.name} className="subject-item">
                <span className="subject-rank">#{index + 1}</span>
                <span className="subject-name">{subject.name}</span>
                <span className="subject-average" style={{ color: '#EF4444' }}>
                  {subject.average}%
                </span>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
};

export default ViewPerformance;