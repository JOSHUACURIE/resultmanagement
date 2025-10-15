import React, { useState, useEffect } from 'react';
import { get } from '../../api/api';
import Card from '../../components/ui/Card';
import Button from '../../components/ui/Button';
import Table from '../../components/ui/Table';
import Select from '../../components/ui/Select';
import Badge from '../../components/ui/Badge';
import './ResultsView.css';

const SubmittedResultsView = () => {
  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [filters, setFilters] = useState({
    term_id: '',
    class_id: '',
    stream_id: '',
    subject_id: ''
  });
  
  const [filterOptions, setFilterOptions] = useState({
    terms: [],
    classes: [],
    streams: [],
    subjects: []
  });

  // Safe comparison function for class_id filtering
  const safeClassIdCompare = (itemClassId, selectedClassId) => {
    if (!itemClassId || !selectedClassId) return false;
    return itemClassId.toString() === selectedClassId.toString();
  };

  // Fetch filter options including subjects
  useEffect(() => {
    const fetchFilterOptions = async () => {
      try {
        const [termsRes, classesRes, streamsRes, subjectsRes] = await Promise.all([
          get('/terms'),
          get('/classes'),
          get('/streams'),
          get('/subjects')
        ]);

        const terms = termsRes?.success ? termsRes.data : (Array.isArray(termsRes) ? termsRes : []);
        const classes = classesRes?.success ? classesRes.data : (Array.isArray(classesRes) ? classesRes : []);
        const streams = streamsRes?.success ? streamsRes.data : (Array.isArray(streamsRes) ? streamsRes : []);
        const subjects = subjectsRes?.success ? subjectsRes.data : (Array.isArray(subjectsRes) ? subjectsRes : []);

        setFilterOptions({
          terms: terms || [],
          classes: (classes || []).filter(cls => cls?.is_active !== false),
          streams: (streams || []).filter(stream => stream?.is_active !== false),
          subjects: (subjects || []).filter(subject => subject?.is_active !== false)
        });
      } catch (err) {
        console.error('Failed to load filter options:', err);
        setError('Failed to load filter options. Please try again.');
      }
    };
    
    fetchFilterOptions();
  }, []);

  // Fetch submitted results
  const fetchResults = async () => {
    if (!filters.term_id || !filters.class_id) {
      setError('Please select term and class');
      return;
    }

    setLoading(true);
    setError('');

    try {
      const selectedTerm = filterOptions.terms.find(term => 
        term?.term_id?.toString() === filters.term_id
      );

      if (!selectedTerm) {
        setError('Selected term not found');
        return;
      }

      const queryParams = new URLSearchParams({
        term_id: filters.term_id,
        class_id: filters.class_id,
        academic_year: selectedTerm.academic_year,
        ...(filters.stream_id && { stream_id: filters.stream_id }),
        ...(filters.subject_id && { subject_id: filters.subject_id })
      });

      const response = await get(`/results/submitted?${queryParams}`);
      
      if (response?.success) {
        setResults(response.data || []);
      } else {
        setError(response?.message || 'Failed to fetch results');
        setResults([]);
      }
    } catch (err) {
      console.error('Failed to fetch submitted results:', err);
      setError('Failed to load results. Please try again.');
      setResults([]);
    } finally {
      setLoading(false);
    }
  };

  const handleFilterChange = (key, value) => {
    setFilters(prev => ({
      ...prev,
      [key]: value
    }));

    if (key === 'class_id') {
      setFilters(prev => ({
        ...prev,
        stream_id: '',
        subject_id: ''
      }));
    }
  };

  const handleClearFilters = () => {
    setFilters({
      term_id: '',
      class_id: '',
      stream_id: '',
      subject_id: ''
    });
    setResults([]);
    setError('');
  };

  const handleExportExcel = async () => {
    if (!filters.term_id || !filters.class_id) {
      setError('Please select term and class');
      return;
    }

    try {
      const selectedTerm = filterOptions.terms.find(term => 
        term?.term_id?.toString() === filters.term_id
      );

      if (!selectedTerm) {
        setError('Selected term not found');
        return;
      }

      const queryParams = new URLSearchParams({
        term_id: filters.term_id,
        class_id: filters.class_id,
        academic_year: selectedTerm.academic_year,
        ...(filters.stream_id && { stream_id: filters.stream_id }),
        ...(filters.subject_id && { subject_id: filters.subject_id })
      });

      window.open(`/api/results/export-all?${queryParams}`, '_blank');
    } catch (err) {
      console.error('Export failed:', err);
      setError('Failed to export results');
    }
  };

  const handleExportStudentResult = async (studentId, admissionNumber) => {
    try {
      const selectedTerm = filterOptions.terms.find(term => 
        term?.term_id?.toString() === filters.term_id
      );

      if (!selectedTerm) {
        setError('Selected term not found');
        return;
      }

      const queryParams = new URLSearchParams({
        term_id: filters.term_id,
        academic_year: selectedTerm.academic_year,
        ...(filters.subject_id && { subject_id: filters.subject_id })
      });

      window.open(`/api/results/export-student/${studentId}?${queryParams}`, '_blank');
    } catch (err) {
      console.error('Export failed:', err);
      setError(`Failed to export result for ${admissionNumber}`);
    }
  };

  const getSelectedTermInfo = () => {
    if (!filters.term_id) return null;
    const selectedTerm = filterOptions.terms.find(term => 
      term?.term_id?.toString() === filters.term_id
    );
    return selectedTerm ? `${selectedTerm.term_name} - ${selectedTerm.academic_year}` : '';
  };

  const calculateGrade = (score) => {
    if (!score && score !== 0) return '-';
    const numScore = parseFloat(score);
    if (isNaN(numScore)) return '-';
    if (numScore >= 80) return 'A';
    if (numScore >= 75) return 'A-';
    if (numScore >= 70) return 'B+';
    if (numScore >= 65) return 'B';
    if (numScore >= 55) return 'B-';
    if (numScore >= 50) return 'C+';
    if (numScore >= 45) return 'C';
    if (numScore >= 40) return 'C-';
    if (numScore >= 35) return 'D+';
    if (numScore >= 30) return 'D';
    if (numScore >= 25) return 'D-';
    return 'E';
  };

  // Dynamic subject columns
  const getSubjectColumns = () => {
    if (results.length === 0) return [];
    
    // Get unique subjects from results
    const subjects = Array.from(
      new Set(
        results.flatMap(student => 
          (student.subject_scores || []).map(score => score?.subject_name).filter(Boolean)
        ).filter(Boolean)
      )
    );

    return subjects.map(subject => ({
      key: `subject_${subject}`,
      header: subject,
      render: (_, row) => {
        const subjectScore = row.subject_scores?.find(
          score => score?.subject_name === subject
        );
        const score = subjectScore ? subjectScore.score : '-';
        const grade = calculateGrade(score);
        return (
          <div className="flex flex-col items-center text-center">
            <span className="font-semibold text-sm">{score}</span>
            <Badge variant={getGradeVariant(grade)} size="sm" className="mt-1">
              {grade}
            </Badge>
          </div>
        );
      }
    }));
  };

  // Base columns
  const baseColumns = [
    { 
      key: 'admission_number', 
      header: 'Adm No.',
      render: (value) => <span className="font-mono text-sm">{value || '-'}</span>
    },
    { 
      key: 'fullname', 
      header: 'Student Name',
      render: (value) => <span className="font-medium">{value || '-'}</span>
    },
    { 
      key: 'class_name', 
      header: 'Class',
      render: (value) => <span>{value || '-'}</span>
    },
    { 
      key: 'stream_name', 
      header: 'Stream',
      render: (value) => <span>{value || '-'}</span>
    }
  ];

  // Summary columns
  const summaryColumns = [
    { 
      key: 'total_score', 
      header: 'Total',
      render: (value) => {
        const num = parseFloat(value);
        return <span className="font-semibold">{isNaN(num) ? '-' : num.toFixed(2)}</span>
      }
    },
    { 
      key: 'average_score', 
      header: 'Average',
      render: (value) => {
        const num = parseFloat(value);
        return <span className="font-semibold">{isNaN(num) ? '-' : num.toFixed(2)}</span>
      }
    },
    { 
      key: 'class_rank', 
      header: 'Rank',
      render: (value) => <span className="font-bold">{value || '-'}</span>
    },
    {
      key: 'actions',
      header: 'Actions',
      render: (_, row) => (
        <Button
          variant="outline"
          size="sm"
          onClick={() => handleExportStudentResult(row.student_id, row.admission_number)}
          className="export-btn"
          disabled={!row.student_id}
        >
          Export
        </Button>
      )
    }
  ];

  const resultColumns = [...baseColumns, ...getSubjectColumns(), ...summaryColumns];

  const getGradeVariant = (grade) => {
    const gradeVariants = {
      'A': 'success',
      'B': 'info', 
      'C': 'warning',
      'D': 'error',
      'E': 'error',
      '-': 'default'
    };
    return gradeVariants[grade?.charAt(0)] || 'default';
  };

  const calculateClassAverage = () => {
    if (results.length === 0) return '0.00';
    const total = results.reduce((sum, student) => {
      const avg = parseFloat(student.average_score || 0);
      return sum + (!isNaN(avg) ? avg : 0);
    }, 0);
    return (total / results.length).toFixed(2);
  };

  const canFetchResults = filters.term_id && filters.class_id;

  return (
    <div className="results-view">
      {/* Header */}
      <div className="results-header">
        <h1 className="results-title">Student Results Dashboard</h1>
        <p className="results-subtitle">Comprehensive view of student academic performance</p>
      </div>

      {/* Filters - Updated to use flex layout */}
      <Card className="filters-card">
        <h2 className="filters-title">Filter Results</h2>
        
        <div className="filters-grid">
          <div className="filter-group">
            <label className="filter-label required">
              Term & Academic Year
            </label>
            <select
              className="native-select"
              value={filters.term_id}
              onChange={(e) => handleFilterChange('term_id', e.target.value)}
              required
            >
              <option value="">Select Term & Academic Year</option>
              {Object.entries(
                filterOptions.terms.reduce((acc, term) => {
                  const year = term?.academic_year;
                  if (year && !acc[year]) acc[year] = [];
                  if (year) acc[year].push(term);
                  return acc;
                }, {})
              )
                .sort(([a], [b]) => b.localeCompare(a))
                .map(([year, terms]) => (
                  <optgroup key={year} label={`Academic Year ${year}`}>
                    {terms
                      ?.sort((a, b) => (a?.term_number || 0) - (b?.term_number || 0))
                      ?.map(term => (
                        <option key={term?.term_id} value={term?.term_id}>
                          {term?.term_name} (Term {term?.term_number})
                        </option>
                      ))
                    }
                  </optgroup>
                ))
              }
            </select>
            {filters.term_id && (
              <div className="selected-term-info">
                Selected: {getSelectedTermInfo()}
              </div>
            )}
          </div>

          <div className="filter-group">
            <label className="filter-label required">
              Class
            </label>
            <select
              className="native-select"
              value={filters.class_id}
              onChange={(e) => handleFilterChange('class_id', e.target.value)}
              required
            >
              <option value="">Select Class</option>
              {filterOptions.classes?.map(cls => (
                <option key={cls?.class_id} value={cls?.class_id}>
                  {cls?.class_name} (Level {cls?.class_level})
                </option>
              ))}
            </select>
          </div>

          <div className="filter-group">
            <label className="filter-label">
              Stream
            </label>
            <select
              className="native-select"
              value={filters.stream_id}
              onChange={(e) => handleFilterChange('stream_id', e.target.value)}
              disabled={!filters.class_id}
            >
              <option value="">All Streams</option>
              {filterOptions.streams
                ?.filter(stream => safeClassIdCompare(stream?.class_id, filters.class_id))
                ?.map(stream => (
                  <option key={stream?.stream_id} value={stream?.stream_id}>
                    {stream?.stream_name}
                  </option>
                ))
              }
            </select>
            {filters.class_id && 
              filterOptions.streams?.filter(stream => safeClassIdCompare(stream?.class_id, filters.class_id)).length === 0 && (
              <div className="no-streams-info">
                No streams available for this class
              </div>
            )}
          </div>

          <div className="filter-group">
            <label className="filter-label">
              Subject
            </label>
            <select
              className="native-select"
              value={filters.subject_id}
              onChange={(e) => handleFilterChange('subject_id', e.target.value)}
              disabled={!filters.class_id}
            >
              <option value="">All Subjects</option>
              {filterOptions.subjects
                ?.filter(subject => safeClassIdCompare(subject?.class_id, filters.class_id))
                ?.map(subject => (
                  <option key={subject?.subject_id} value={subject?.subject_id}>
                    {subject?.subject_name}
                  </option>
                ))
              }
            </select>
          </div>
        </div>

        {/* Filter Actions */}
        <div className="filter-actions">
          <button
            className="btn btn-primary"
            onClick={fetchResults}
            disabled={loading || !canFetchResults}
          >
            {loading ? 'Loading Results...' : 'View Results'}
          </button>
          
          <button
            className="btn btn-outline"
            onClick={handleClearFilters}
            disabled={loading}
          >
            Clear Filters
          </button>
          
          <button
            className="btn btn-outline"
            onClick={handleExportExcel}
            disabled={results.length === 0 || loading || !canFetchResults}
          >
            Export to Excel
          </button>
        </div>
        
        <div className="required-fields-note">
          * Required fields
        </div>
      </Card>

      {/* Error Message */}
      {error && (
        <Card className="error-card">
          <div className="error-message">
            <span>{error}</span>
            <button 
              className="btn btn-text"
              onClick={() => setError('')}
            >
              Dismiss
            </button>
          </div>
        </Card>
      )}

      {/* Results Summary - Updated to use flex layout */}
      {results.length > 0 && (
        <Card className="results-card">
          <div className="results-summary">
            <h3 className="summary-title">
              Results Summary ({results.length} students)
            </h3>
            <div className="summary-flex">
              <div className="summary-stat">
                <span className="stat-label">Academic Period</span>
                <span className="stat-value">{getSelectedTermInfo()}</span>
              </div>
              <div className="summary-stat">
                <span className="stat-label">Top Student</span>
                <span className="stat-value">
                  {results[0]?.fullname || 'N/A'} ({parseFloat(results[0]?.average_score || 0).toFixed(2)})
                </span>
              </div>
              <div className="summary-stat">
                <span className="stat-label">Class Average</span>
                <span className="stat-value highlight">{calculateClassAverage()}</span>
              </div>
              <div className="summary-stat">
                <span className="stat-label">Total Subjects</span>
                <span className="stat-value">{getSubjectColumns().length}</span>
              </div>
            </div>
          </div>

          {/* Results Table */}
          <div className="results-table-container">
            <Table
              columns={resultColumns}
              data={results}
              className="results-table"
            />
          </div>
        </Card>
      )}

      {/* Loading State */}
      {loading && (
        <Card className="loading-card">
          <div className="loading-state">
            <div className="loading-spinner"></div>
            <p>Loading results...</p>
          </div>
        </Card>
      )}

      {/* Initial State */}
      {!loading && results.length === 0 && !canFetchResults && (
        <Card className="empty-card">
          <div className="empty-state">
            <h3>Select Filters to View Results</h3>
            <p>Please select term and class to view submitted results.</p>
          </div>
        </Card>
      )}

      {/* Empty State */}
      {!loading && results.length === 0 && canFetchResults && (
        <Card className="empty-card">
          <div className="empty-state">
            <h3>No Results Found</h3>
            <p>No submitted results found for the selected criteria.</p>
            <button 
              className="btn btn-outline"
              onClick={fetchResults}
            >
              Try Again
            </button>
          </div>
        </Card>
      )}
    </div>
  );
};

export default SubmittedResultsView;