import React, { useState, useEffect } from 'react';
import { get } from '../../api/api';
import Card from '../../components/ui/Card';
import Button from '../../components/ui/Button';
import Modal from '../../components/ui/Modal';
import Select from '../../components/ui/Select';
import Badge from '../../components/ui/Badge';
import { 
  exportResultsToExcel, 
  exportIndividualResultToWord, 
  exportIndividualResultAsHTML,
  downloadBlob, 
  generateFilename,
  calculateGrade,
  calculateTotalGrade,
  getUniqueSubjects,
  loadImageAsBase64
} from '../../utils/exportUtils';
import './AnalyzeResults.css';

const AnalyzeResults = () => {
  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [exportingGeneral, setExportingGeneral] = useState(false);
  const [exportingAllIndividual, setExportingAllIndividual] = useState(false);
  const [exportingIndividual, setExportingIndividual] = useState(false);
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
  const [selectedStudent, setSelectedStudent] = useState(null);
  const [exportProgress, setExportProgress] = useState({ current: 0, total: 0 });

  // Fetch filter options
  useEffect(() => {
    const fetchFilterOptions = async () => {
      try {
        const [termsRes, classesRes, streamsRes, subjectsRes] = await Promise.all([
          get('/terms'),
          get('/classes'),
          get('/streams'),
          get('/subjects')
        ]);

        const terms = termsRes.success ? termsRes.data : (Array.isArray(termsRes) ? termsRes : []);
        const classes = classesRes.success ? classesRes.data : (Array.isArray(classesRes) ? classesRes : []);
        const streams = streamsRes.success ? streamsRes.data : (Array.isArray(streamsRes) ? streamsRes : []);
        const subjects = subjectsRes.success ? subjectsRes.data : (Array.isArray(subjectsRes) ? subjectsRes : []);

        setFilterOptions({
          terms: terms || [],
          classes: (classes || []).filter(cls => cls.is_active !== false),
          streams: (streams || []).filter(stream => 
            stream.is_active !== false && 
            stream.class_id !== undefined && 
            stream.class_id !== null
          ),
          subjects: (subjects || []).filter(subject => 
            subject.is_active !== false && 
            subject.class_id !== undefined && 
            subject.class_id !== null
          )
        });
      } catch (err) {
        console.error('Failed to load filter options:', err);
        setError('Failed to load filter options. Please try again.');
      }
    };
    
    fetchFilterOptions();
  }, []);

  // Fetch results
  const fetchResults = async () => {
    if (!filters.term_id || !filters.class_id) {
      setError('Please select term and class');
      return;
    }

    setLoading(true);
    setError('');

    try {
      const selectedTerm = filterOptions.terms.find(term => 
        term.term_id.toString() === filters.term_id
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
      
      if (response.success) {
        const enhancedResults = (response.data || []).map(student => ({
          ...student,
          subject_scores: (student.subject_scores || []).map(subject => ({
            ...subject,
            grade: calculateGrade(subject.score),
            remarks: getGradeRemarks(calculateGrade(subject.score))
          }))
        }));
        setResults(enhancedResults);
      } else {
        setError(response.message || 'Failed to fetch results');
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

  // Handle filter changes
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

  // Clear filters
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

  // Export general results to Excel
  const handleExportGeneralResults = async () => {
    if (!filters.term_id || !filters.class_id) {
      setError('Please select term and class');
      return;
    }

    if (results.length === 0) {
      setError('No results to export');
      return;
    }

    setExportingGeneral(true);
    try {
      const blob = await exportResultsToExcel(results, filters, filterOptions);
      const filename = generateFilename(filters, filterOptions);
      downloadBlob(blob, filename);
      alert(`✅ Successfully downloaded Excel report for ${results.length} students!`);
    } catch (err) {
      console.error('Export failed:', err);
      setError('Failed to export results: ' + err.message);
    } finally {
      setExportingGeneral(false);
    }
  };

  // Export ALL individual results to Word documents with logo
  const handleExportAllIndividualResults = async () => {
    if (!filters.term_id || !filters.class_id) {
      setError('Please select term and class');
      return;
    }

    if (results.length === 0) {
      setError('No results to export');
      return;
    }

    setExportingAllIndividual(true);
    setExportProgress({ current: 0, total: results.length });
    
    try {
      let successCount = 0;
      let errorCount = 0;

      // Load logo image as base64 for Word documents
      const logoUrl = '/logo.jpg';
      const imageBase64 = await loadImageAsBase64(logoUrl);

      const processStudentsSequentially = async () => {
        for (let i = 0; i < results.length; i++) {
          const student = results[i];
          try {
            setExportProgress({ current: i + 1, total: results.length });
            
            const subjects = (student.subject_scores || []).map(subject => ({
              subject_name: subject.subject_name,
              score: subject.score,
              grade: subject.grade || calculateGrade(subject.score),
              remarks: subject.remarks || getGradeRemarks(subject.grade || calculateGrade(subject.score))
            }));
            
            const comments = {
              principal: student.principal_comment || getDefaultPrincipalComment(student),
              class_teacher: student.class_teacher_comment || getDefaultClassTeacherComment(student)
            };

            const blob = await exportIndividualResultToWord(student, subjects, comments, imageBase64);
            const filename = generateFilename(filters, filterOptions, student);
            
            downloadBlob(blob, filename);
            successCount++;
            
            if (i < results.length - 1) {
              await new Promise(resolve => setTimeout(resolve, 300));
            }
          } catch (err) {
            console.error(`Failed to export result for ${student.admission_number}:`, err);
            errorCount++;
          }
        }
      };

      await processStudentsSequentially();

      if (errorCount === 0) {
        alert(`✅ Successfully downloaded ${successCount} individual student reports!`);
      } else {
        alert(`📊 Download completed with ${successCount} successful and ${errorCount} failed downloads.`);
      }
    } catch (err) {
      console.error('Bulk export failed:', err);
      setError('Failed to export individual results: ' + err.message);
    } finally {
      setExportingAllIndividual(false);
      setExportProgress({ current: 0, total: 0 });
    }
  };

  // Export individual student result as Word
  const handleExportIndividualWord = async (student) => {
    setExportingIndividual(true);
    try {
      const subjects = (student.subject_scores || []).map(subject => ({
        subject_name: subject.subject_name,
        score: subject.score,
        grade: subject.grade || calculateGrade(subject.score),
        remarks: subject.remarks || getGradeRemarks(subject.grade || calculateGrade(subject.score))
      }));
      
      const comments = {
        principal: student.principal_comment || getDefaultPrincipalComment(student),
        class_teacher: student.class_teacher_comment || getDefaultClassTeacherComment(student)
      };

      const logoUrl = '/logo.jpg';
      const imageBase64 = await loadImageAsBase64(logoUrl);

      const blob = await exportIndividualResultToWord(student, subjects, comments, imageBase64);
      const filename = generateFilename(filters, filterOptions, student);
      
      downloadBlob(blob, filename);
      alert(`✅ Successfully downloaded Word report for ${student.fullname}!`);
    } catch (err) {
      console.error(`Failed to export Word for ${student.admission_number}:`, err);
      setError('Failed to export Word report: ' + err.message);
    } finally {
      setExportingIndividual(false);
    }
  };

  // Export individual student result as HTML
  const handleExportIndividualHTML = async (student) => {
    setExportingIndividual(true);
    try {
      const subjects = (student.subject_scores || []).map(subject => ({
        subject_name: subject.subject_name,
        score: subject.score,
        grade: subject.grade || calculateGrade(subject.score),
        remarks: subject.remarks || getGradeRemarks(subject.grade || calculateGrade(subject.score))
      }));
      
      const comments = {
        principal: student.principal_comment || getDefaultPrincipalComment(student),
        class_teacher: student.class_teacher_comment || getDefaultClassTeacherComment(student)
      };

      const logoUrl = '/logo.jpg';
      const blob = await exportIndividualResultAsHTML(student, subjects, comments, logoUrl);
      const filename = `Academic_Report_${student.admission_number}_${(student.fullname || student.name).replace(/\s+/g, '_')}.html`;
      
      downloadBlob(blob, filename);
      alert(`✅ Successfully downloaded HTML report for ${student.fullname}!`);
    } catch (err) {
      console.error(`Failed to export HTML for ${student.admission_number}:`, err);
      setError('Failed to export HTML report: ' + err.message);
    } finally {
      setExportingIndividual(false);
    }
  };

  // Get default principal comment based on performance
  const getDefaultPrincipalComment = (student) => {
    const totalMarks = Math.round(student.total_score || 0);
    const percentage = (totalMarks / 1100) * 100;
    
    if (percentage >= 80) return 'Exceptional performance! Maintain this outstanding achievement.';
    if (percentage >= 70) return 'Excellent work! Continue with the same dedication and focus.';
    if (percentage >= 60) return 'Good performance. Keep working hard for even better results.';
    if (percentage >= 50) return 'Satisfactory performance. Focus on improvement in weaker areas.';
    if (percentage >= 40) return 'Needs improvement. Dedicate more time to studies and seek help where needed.';
    return 'Requires immediate attention and academic support.';
  };

  // Get default class teacher comment based on performance
  const getDefaultClassTeacherComment = (student) => {
    const totalMarks = Math.round(student.total_score || 0);
    const percentage = (totalMarks / 1100) * 100;
    const grade = calculateTotalGrade(totalMarks);
    
    if (grade === 'A' || grade === 'A-') return 'Outstanding student with excellent academic discipline and consistent performance.';
    if (grade === 'B+' || grade === 'B') return 'Hardworking student showing great potential and steady improvement.';
    if (grade === 'B-' || grade === 'C+') return 'Good effort shown. Focus on consistency across all subjects.';
    if (grade === 'C' || grade === 'C-') return 'Shows potential but needs to improve study habits and subject mastery.';
    if (grade === 'D+' || grade === 'D') return 'Requires more dedication and regular study routine to improve performance.';
    return 'Needs urgent academic intervention and parental support for improvement.';
  };

  // View student details
  const handleViewStudent = (student) => {
    setSelectedStudent(student);
  };

  // Get grade remarks
  const getGradeRemarks = (grade) => {
    const remarks = {
      'A': 'Excellent - Outstanding performance',
      'A-': 'Excellent - Very strong performance',
      'B+': 'Very Good - Strong performance',
      'B': 'Good - Above average performance',
      'B-': 'Good - Good performance',
      'C+': 'Average - Satisfactory performance',
      'C': 'Below Average - Needs improvement',
      'C-': 'Below Average - Requires attention',
      'D+': 'Weak - Significant improvement needed',
      'D': 'Weak - Much improvement required',
      'D-': 'Weak - Serious attention needed',
      'E': 'Weak- Immediate intervention required'
    };
    return remarks[grade] || 'Performance assessment pending';
  };

  // Get selected term info
  const getSelectedTermInfo = () => {
    if (!filters.term_id) return null;
    const selectedTerm = filterOptions.terms.find(term => 
      term.term_id.toString() === filters.term_id
    );
    return selectedTerm ? `${selectedTerm.term_name} - ${selectedTerm.academic_year}` : '';
  };

  // Calculate class average
  const calculateClassAverage = () => {
    if (results.length === 0) return '0.00';
    const total = results.reduce((sum, student) => sum + parseFloat(student.average_score || 0), 0);
    return (total / results.length ).toFixed(2);
  };

  // Grade variant for badges
  const getGradeVariant = (grade) => {
    const gradeVariants = {
      'A': 'success',
      'B': 'info',
      'C': 'warning',
      'D': 'error',
      'E': 'error',
      '-': 'default'
    };
    return gradeVariants[grade?.[0]] || 'default';
  };

  // Get subject abbreviation for table display
  const getSubjectAbbreviation = (subjectName) => {
    const abbreviations = {
      'ENGLISH': 'ENG',
      'MATHEMATICS': 'MATH',
      'KISWAHILI': 'KISW',
      'BIOLOGY': 'BIO',
      'PHYSICS': 'PHY',
      'CHEMISTRY': 'CHEM',
      'GEOGRAPHY': 'GEO',
      'HISTORY': 'HIST',
      'CRE': 'CRE',
      'IRE': 'IRE',
      'HRE': 'HRE',
      'BUSINESS': 'BST',
    };

    const upperSubject = subjectName.toUpperCase();
    return abbreviations[upperSubject] || subjectName.substring(0, 4).toUpperCase();
  };

  // Dynamic subject columns with score + grade format
  const getSubjectColumns = () => {
    if (results.length === 0) return [];
    
    const subjects = getUniqueSubjects(results);

    return subjects.map(subject => ({
      key: `subject_${subject}`,
      header: getSubjectAbbreviation(subject),
      render: (_, row) => {
        const subjectScore = row.subject_scores?.find(
          score => score.subject_name === subject
        );
        const score = subjectScore ? subjectScore.score : '-';
        if (score === '-') {
          return <span className="score-value">-</span>;
        }
        
        const grade = subjectScore?.grade || calculateGrade(score);
        return (
          <div className="subject-score-with-grade">
            <span className="score-value">{Math.round(score)}</span>
            <Badge variant={getGradeVariant(grade)} size="sm">
              {grade}
            </Badge>
          </div>
        );
      }
    }));
  };

  // Desktop table columns
  const baseColumns = [
    { 
      key: 'index', 
      header: '#',
      render: (_, row, index) => <span className="position-number">{index + 1}</span>
    },
    { 
      key: 'admission_number', 
      header: 'ADM NO',
      render: (value) => <span className="admission-number">{value}</span>
    },
    { 
      key: 'fullname', 
      header: 'NAME',
      render: (value) => <span className="student-name">{value}</span>,
      onClick: (row) => handleViewStudent(row)
    }
  ];

  const summaryColumns = [
    { 
      key: 'total_score', 
      header: 'TOTAL',
      render: (value, row) => {
        const totalMarks = Math.round(value);
        const grade = calculateTotalGrade(totalMarks);
        return (
          <div className="total-marks-with-grade">
            <span className="total-score">{totalMarks}</span>
            <Badge variant={getGradeVariant(grade)}>
              {grade}
            </Badge>
          </div>
        );
      }
    },
    { 
      key: 'grade', 
      header: 'GRADE',
      render: (value, row) => {
        const totalMarks = Math.round(row.total_score);
        const grade = calculateTotalGrade(totalMarks);
        return (
          <Badge variant={getGradeVariant(grade)}>
            {grade}
          </Badge>
        );
      }
    },
    { 
      key: 'class_rank', 
      header: 'RANK',
      render: (value) => <span className="rank-value">{value}</span>
    }
  ];

  const desktopColumns = [...baseColumns, ...getSubjectColumns(), ...summaryColumns];

  const canFetchResults = filters.term_id && filters.class_id;

  return (
    <div className="analyze-results">
      {/* Header */}
      <div className="page-header">
        <div className="header-content">
          <h1 className="page-title">Analyze Results</h1>
          <p className="page-subtitle">View, download, and analyze student academic performance</p>
        </div>
        <div className="header-actions">
          <Button
            variant="primary"
            onClick={handleExportGeneralResults}
            disabled={results.length === 0 || exportingGeneral || !canFetchResults}
            className="export-general-btn"
          >
            {exportingGeneral ? 'Exporting...' : 'Export General Results (Excel)'}
          </Button>
          <Button
            variant="success"
            onClick={handleExportAllIndividualResults}
            disabled={results.length === 0 || exportingAllIndividual || !canFetchResults}
            className="export-all-individual-btn"
          >
            {exportingAllIndividual ? (
              <span className="export-progress">
                Exporting... {exportProgress.current}/{exportProgress.total}
              </span>
            ) : (
              'Download All Individual Reports'
            )}
          </Button>
        </div>
      </div>

      {/* Filters */}
      <Card className="filters-card">
        <div className="card-content">
          <h2 className="filters-title">Filter Results</h2>
          <div className="filters-grid">
            <div className="filter-group">
              <Select
                label="Term & Academic Year *"
                name="term_id"
                value={filters.term_id}
                onChange={(e) => handleFilterChange('term_id', e.target.value)}
                required={true}
                className="filter-select"
              >
                <option value="">Select Term & Academic Year</option>
                {Object.entries(
                  filterOptions.terms.reduce((acc, term) => {
                    const year = term.academic_year;
                    if (!acc[year]) acc[year] = [];
                    acc[year].push(term);
                    return acc;
                  }, {})
                )
                  .sort(([a], [b]) => b.localeCompare(a))
                  .map(([year, terms]) => (
                    <optgroup key={year} label={`Academic Year ${year}`}>
                      {terms
                        .sort((a, b) => a.term_number - b.term_number)
                        .map(term => (
                          <option key={term.term_id} value={term.term_id}>
                            {term.term_name} (Term {term.term_number})
                          </option>
                        ))
                      }
                    </optgroup>
                  ))
                }
              </Select>
              {filters.term_id && (
                <div className="selected-term-info">
                  Selected: {getSelectedTermInfo()}
                </div>
              )}
            </div>

            <div className="filter-group">
              <Select
                label="Class *"
                name="class_id"
                value={filters.class_id}
                onChange={(e) => handleFilterChange('class_id', e.target.value)}
                required={true}
                className="filter-select"
              >
                <option value="">Select Class</option>
                {filterOptions.classes.map(cls => (
                  <option key={cls.class_id} value={cls.class_id}>
                    {cls.class_name} (Level {cls.class_level})
                  </option>
                ))}
              </Select>
            </div>

            <div className="filter-group">
              <Select
                label="Stream"
                name="stream_id"
                value={filters.stream_id}
                onChange={(e) => handleFilterChange('stream_id', e.target.value)}
                disabled={!filters.class_id}
                className="filter-select"
              >
                <option value="">All Streams</option>
                {filterOptions.streams
                  .filter(stream => 
                    stream.class_id?.toString() === filters.class_id
                  )
                  .map(stream => (
                    <option key={stream.stream_id} value={stream.stream_id}>
                      {stream.stream_name}
                    </option>
                  ))
                }
              </Select>
              {filters.class_id && filterOptions.streams.filter(stream => 
                stream.class_id?.toString() === filters.class_id
              ).length === 0 && (
                <div className="no-streams-info">
                  No streams available for this class
                </div>
              )}
            </div>

            <div className="filter-group">
              <Select
                label="Subject"
                name="subject_id"
                value={filters.subject_id}
                onChange={(e) => handleFilterChange('subject_id', e.target.value)}
                disabled={!filters.class_id}
                className="filter-select"
              >
                <option value="">All Subjects</option>
                {filterOptions.subjects
                  .filter(subject => 
                    subject.class_id?.toString() === filters.class_id
                  )
                  .map(subject => (
                    <option key={subject.subject_id} value={subject.subject_id}>
                      {subject.subject_name}
                    </option>
                  ))
                }
              </Select>
            </div>
          </div>

          <div className="filter-actions">
            <Button
              variant="primary"
              onClick={fetchResults}
              disabled={loading || !canFetchResults}
              className="view-results-btn"
            >
              {loading ? 'Loading Results...' : 'View Results'}
            </Button>
            
            <Button
              variant="outline"
              onClick={handleClearFilters}
              disabled={loading}
              className="clear-filters-btn"
            >
              Clear Filters
            </Button>
          </div>
          
          <div className="required-fields-note">
            * Required fields
          </div>
        </div>
      </Card>

      {/* Error Message */}
      {error && (
        <Card className="error-card">
          <div className="error-content">
            <span className="error-message">{error}</span>
            <Button 
              variant="text" 
              size="sm" 
              onClick={() => setError('')}
              className="dismiss-btn"
            >
              Dismiss
            </Button>
          </div>
        </Card>
      )}

      {/* Results Section */}
      {results.length > 0 && (
        <Card className="results-card">
          <div className="results-content">
            <div className="results-header">
              <h3 className="results-title">
                Results Summary ({results.length} students)
              </h3>
              <div className="results-header-actions">
                <div className="class-average">
                  Class Average: <span className="average-value">{calculateClassAverage()}</span>
                </div>
                <div className="export-buttons">
                  <Button
                    variant="outline"
                    onClick={handleExportGeneralResults}
                    disabled={exportingGeneral}
                    className="export-results-btn"
                  >
                    {exportingGeneral ? 'Exporting...' : 'Export Results (Excel)'}
                  </Button>
                  <Button
                    variant="success"
                    onClick={handleExportAllIndividualResults}
                    disabled={exportingAllIndividual}
                    className="export-all-individual-btn"
                  >
                    {exportingAllIndividual ? (
                      <span className="export-progress">
                        Downloading... {exportProgress.current}/{exportProgress.total}
                      </span>
                    ) : (
                      `Download All (${results.length})`
                    )}
                  </Button>
                </div>
              </div>
            </div>
            
            <div className="summary-grid">
              <div className="summary-item academic-period">
                <span className="summary-label">Academic Period</span>
                <span className="summary-value">{getSelectedTermInfo()}</span>
              </div>
              <div className="summary-item top-student">
                <span className="summary-label">Top Student</span>
                <span className="summary-value">
                  {results[0]?.fullname} ({parseFloat(results[0]?.average_score).toFixed(2)})
                </span>
              </div>
              <div className="summary-item total-subjects">
                <span className="summary-label">Total Subjects</span>
                <span className="summary-value">{getUniqueSubjects(results).length}</span>
              </div>
              <div className="summary-item performance">
                <span className="summary-label">Performance</span>
                <span className="summary-value">
                  {calculateClassAverage() >= 70 ? 'Excellent' : 
                   calculateClassAverage() >= 60 ? 'Good' : 
                   calculateClassAverage() >= 50 ? 'Average' : 'Needs Improvement'}
                </span>
              </div>
            </div>

            {/* Desktop Table */}
            <div className="table-container desktop-only">
              <table className="results-table desktop-table">
                <thead>
                  <tr>
                    {desktopColumns.map(column => (
                      <th key={column.key}>{column.header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {results.map((student, index) => (
                    <tr 
                      key={student.student_id || index} 
                      className="student-row"
                      onClick={() => baseColumns.find(col => col.key === 'fullname')?.onClick?.(student)}
                    >
                      {desktopColumns.map(column => (
                        <td key={column.key}>
                          {column.render ? column.render(student[column.key], student, index) : student[column.key]}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Mobile Cards */}
            <div className="mobile-students-list mobile-only">
              {results.map((student, index) => (
                <div 
                  key={student.student_id || index} 
                  className="student-card"
                  onClick={() => handleViewStudent(student)}
                >
                  <div className="student-card-header">
                    <div className="student-card-main">
                      <div className="student-card-name">{student.fullname}</div>
                      <div className="student-card-admission">{student.admission_number}</div>
                      <div className="student-card-rank">Rank: {index + 1}</div>
                    </div>
                    <div className="student-card-overall">
                      <div className="total-score-mobile">
                        Total: <strong>{Math.round(student.total_score)}</strong>
                      </div>
                      <Badge variant={getGradeVariant(calculateTotalGrade(student.total_score))}>
                        {calculateTotalGrade(student.total_score)}
                      </Badge>
                    </div>
                  </div>
                  
                  <div className="student-card-scores">
                    <div className="scores-header">Subject Scores</div>
                    <div className="scores-grid">
                      {(student.subject_scores || []).slice(0, 4).map((subject, idx) => (
                        <div key={idx} className="score-item">
                          <span className="subject-abbr">{getSubjectAbbreviation(subject.subject_name)}</span>
                          <span className="score-value">{Math.round(subject.score)}</span>
                          <Badge variant={getGradeVariant(subject.grade)} size="sm">
                            {subject.grade}
                          </Badge>
                        </div>
                      ))}
                      {(student.subject_scores || []).length > 4 && (
                        <div className="more-scores">
                          +{(student.subject_scores || []).length - 4} more subjects
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="student-card-details">
                    <div className="detail-item">
                      <span className="detail-label">Average:</span>
                      <span className="detail-value">{parseFloat(student.average_score).toFixed(1)}</span>
                    </div>
                    <div className="detail-item">
                      <span className="detail-label">Class Rank:</span>
                      <span className="detail-value">{student.class_rank}</span>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </Card>
      )}

      {/* Loading State */}
      {loading && (
        <Card className="loading-card">
          <div className="loading-content">
            <div className="loading-spinner"></div>
            <p className="loading-text">Loading results...</p>
          </div>
        </Card>
      )}

      {/* Initial State */}
      {!loading && results.length === 0 && !canFetchResults && (
        <Card className="empty-card">
          <div className="empty-content">
            <h3 className="empty-title">Select Filters to View Results</h3>
            <p className="empty-description">Please select term and class to view submitted results.</p>
          </div>
        </Card>
      )}

      {/* Empty State */}
      {!loading && results.length === 0 && canFetchResults && (
        <Card className="empty-card">
          <div className="empty-content">
            <h3 className="empty-title">No Results Found</h3>
            <p className="empty-description">No submitted results found for the selected criteria.</p>
            <Button 
              variant="outline" 
              onClick={fetchResults}
              className="try-again-btn"
            >
              Try Again
            </Button>
          </div>
        </Card>
      )}

      {/* Student Details Modal */}
      <Modal
        isOpen={!!selectedStudent}
        onClose={() => setSelectedStudent(null)}
        title="Student Results Details"
        size="lg"
      >
        {selectedStudent && (
          <div className="student-details-modal">
            <div className="student-basic-info">
              <div className="info-item">
                <strong>Name:</strong> {selectedStudent.fullname}
              </div>
              <div className="info-item">
                <strong>Admission No:</strong> {selectedStudent.admission_number}
              </div>
              <div className="info-item">
                <strong>Total Marks:</strong> {Math.round(selectedStudent.total_score)}
              </div>
              <div className="info-item">
                <strong>Average:</strong> {parseFloat(selectedStudent.average_score).toFixed(2)}
              </div>
              <div className="info-item">
                <strong>Overall Grade:</strong>
                <Badge variant={getGradeVariant(calculateTotalGrade(selectedStudent.total_score))}>
                  {calculateTotalGrade(selectedStudent.total_score)}
                </Badge>
              </div>
              <div className="info-item">
                <strong>Class Rank:</strong> {selectedStudent.class_rank}
              </div>
            </div>

            <div className="subject-scores-details">
              <h4>Subject Scores</h4>
              <div className="scores-table">
                {(selectedStudent.subject_scores || []).map((subject, index) => (
                  <div key={index} className="score-row">
                    <div className="subject-name">{subject.subject_name}</div>
                    <div className="score-details">
                      <span className="score">{Math.round(subject.score)}</span>
                      <Badge variant={getGradeVariant(subject.grade)} size="sm">
                        {subject.grade}
                      </Badge>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="modal-actions">
              <Button
                variant="primary"
                onClick={() => handleExportIndividualWord(selectedStudent)}
                disabled={exportingIndividual}
                className="export-word-btn"
              >
                {exportingIndividual ? 'Exporting...' : 'Export as Word'}
              </Button>
              <Button
                variant="outline"
                onClick={() => handleExportIndividualHTML(selectedStudent)}
                disabled={exportingIndividual}
                className="export-html-btn"
              >
                {exportingIndividual ? 'Exporting...' : 'Export as HTML'}
              </Button>
              <Button
                variant="text"
                onClick={() => setSelectedStudent(null)}
                className="close-btn"
              >
                Close
              </Button>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
};

export default AnalyzeResults;
