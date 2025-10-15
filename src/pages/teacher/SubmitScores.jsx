import React, { useState, useEffect } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { post, get } from '../../api/api';
import { teacherApi } from '../../api/teacherApi';
import './SubmitScores.css';

const SubmitScores = () => {
  const navigate = useNavigate();
  const location = useLocation();

  const urlParams = new URLSearchParams(location.search);
  const initialAssignmentId = urlParams.get('assignmentId');

  const [assignments, setAssignments] = useState([]);
  const [students, setStudents] = useState([]);
  const [selectedAssignment, setSelectedAssignment] = useState(initialAssignmentId || '');
  const [scores, setScores] = useState({});
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [successMessage, setSuccessMessage] = useState('');
  const [activeTerm, setActiveTerm] = useState(null);
  const [academicYear, setAcademicYear] = useState('2025');
  const [availableTerms, setAvailableTerms] = useState([]);
  const [selectedTermId, setSelectedTermId] = useState('3');

  // Fetch available terms
  const fetchTerms = async () => {
    try {
      const response = await get('/terms?is_active=true');
      console.log('Terms response:', response);

      if (response.success) {
        const termsData = response.data || [];
        setAvailableTerms(termsData);

        const term3_2025 = termsData.find((t) => 
          t.term_id == 3 && t.academic_year == '2025'
        );

        if (term3_2025) {
          setActiveTerm(term3_2025);
          setSelectedTermId(term3_2025.term_id.toString());
        } else {
          const activeTermData = termsData.find((t) => t.is_active);
          if (activeTermData) {
            setActiveTerm(activeTermData);
            setSelectedTermId(activeTermData.term_id.toString());
          } else if (termsData.length > 0) {
            setSelectedTermId(termsData[0].term_id.toString());
          } else {
            setError('No active terms available. Please contact the administration.');
          }
        }
      } else {
        setError('Failed to fetch terms');
      }
    } catch (err) {
      console.error('Error fetching terms:', err);
      setError('Failed to load terms: ' + err.message);
    }
  };

  // Fetch teacher's assignments
  const fetchAssignments = async () => {
    try {
      setLoading(true);
      setError('');
      setSuccessMessage('');

      if (!academicYear || !selectedTermId) {
        setError('Please select both academic year and term');
        setLoading(false);
        return;
      }

      const response = await teacherApi.getMyAssignments(academicYear, selectedTermId);
      console.log('Assignments API response:', response);

      if (response && response.success) {
        const assignmentsData = response.data || [];
        setAssignments(assignmentsData);

        if (initialAssignmentId && assignmentsData.find((a) => a.assignment_id == initialAssignmentId)) {
          setSelectedAssignment(initialAssignmentId);
        } else if (assignmentsData.length > 0) {
          setSelectedAssignment(assignmentsData[0].assignment_id.toString());
        } else {
          setError(`You have no teaching assignments for ${academicYear} - Term ${selectedTermId}.`);
        }
      } else {
        throw new Error(response?.message || 'Failed to fetch assignments');
      }
    } catch (err) {
      console.error('Error in fetchAssignments:', err);
      setError(`Failed to load assignments: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  // Fetch students for selected assignment
  const fetchStudents = async (assignmentId) => {
    if (!assignmentId) {
      setStudents([]);
      setScores({});
      return;
    }

    try {
      setLoading(true);
      setError('');
      setSuccessMessage('');

      const assignmentResponse = await get(`/assignments/${assignmentId}`);
      console.log('Assignment details:', assignmentResponse);

      if (!assignmentResponse.success) {
        throw new Error(assignmentResponse.message || 'Failed to fetch assignment details');
      }

      const assignment = assignmentResponse.data;
      const classId = assignment.class_id;
      const streamId = assignment.stream_id;

      if (!classId) {
        throw new Error('Assignment does not have a valid class');
      }

      const params = new URLSearchParams({
        class_id: classId,
        is_active: 'true',
      });

      if (streamId) {
        params.append('stream_id', streamId);
      }

      const studentsResponse = await get(`/students?${params.toString()}`);
      console.log('Students response:', studentsResponse);

      if (studentsResponse.success) {
        const studentsData = studentsResponse.data || [];
        setStudents(studentsData);

        // Initialize empty scores
        const initialScores = {};
        studentsData.forEach((student) => {
          initialScores[student.student_id] = '';
        });
        setScores(initialScores);
      } else {
        throw new Error(studentsResponse.message || 'Failed to fetch students');
      }
    } catch (err) {
      console.error('Error fetching students:', err);
      setError(`Failed to load students: ${err.message}`);
      setStudents([]);
      setScores({});
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchTerms();
  }, []);

  useEffect(() => {
    if (academicYear && selectedTermId) {
      fetchAssignments();
    }
  }, [academicYear, selectedTermId]);

  useEffect(() => {
    if (selectedAssignment) {
      fetchStudents(selectedAssignment);
    } else {
      setStudents([]);
      setScores({});
    }
  }, [selectedAssignment]);

  const handleScoreChange = (studentId, value) => {
    if (value === '' || (/^\d*\.?\d*$/.test(value) && parseFloat(value) >= 0 && parseFloat(value) <= 100)) {
      setScores((prev) => ({
        ...prev,
        [studentId]: value,
      }));
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setSuccessMessage('');

    const assignmentId = parseInt(selectedAssignment);
    const termId = parseInt(selectedTermId);

    if (!assignmentId || isNaN(assignmentId)) {
      setError('Please select a valid assignment');
      return;
    }

    if (!termId || isNaN(termId)) {
      setError('Please select a valid term');
      return;
    }

    const scoreEntries = Object.entries(scores)
      .filter(([_, score]) => score !== '' && score !== null)
      .map(([studentId, score]) => ({
        student_id: parseInt(studentId),
        score: parseFloat(score),
      }));

    if (scoreEntries.length === 0) {
      setError('Please enter at least one score');
      return;
    }

    const invalidScores = scoreEntries.filter(
      (entry) => isNaN(entry.score) || entry.score < 0 || entry.score > 100
    );

    if (invalidScores.length > 0) {
      setError('All scores must be valid numbers between 0 and 100');
      return;
    }

    setSubmitting(true);

    try {
      const scoreData = {
        assignment_id: assignmentId,
        term_id: termId,
        scores: scoreEntries,
      };

      console.log('Submitting scores:', scoreData);
      const response = await post('/scores/submit', scoreData);

      if (response.success) {
        setSuccessMessage(response.message || `✅ Successfully submitted ${scoreEntries.length} scores!`);
        setScores({});
        setSelectedAssignment('');
        setStudents([]);
        navigate('/teacher/assigned-subjects');
      } else {
        throw new Error(response.message || 'Failed to submit scores');
      }
    } catch (err) {
      console.error('Error submitting scores:', err);
      const errorMessage = err.response?.data?.message || err.message || 'Failed to submit scores. Please try again.';
      setError(errorMessage);
    } finally {
      setSubmitting(false);
    }
  };

  const getAssignmentDisplayName = (assignment) => {
    if (!assignment) return '';

    const subject = assignment.assignment_subject;
    const classInfo = assignment.assignment_class;
    const stream = assignment.assignment_stream;

    return `${subject?.subject_name || 'Unknown'} - ${classInfo?.class_name || 'Unknown'}${
      stream ? ` (${stream.stream_name})` : ''
    }`;
  };

  const getTermName = (termId) => {
    const term = availableTerms.find((t) => t.term_id == termId);
    return term ? term.term_name : `Term ${termId}`;
  };

  const getScoreEmoji = (score) => {
    if (!score) return '';
    const numScore = parseFloat(score);
    if (numScore >= 85) return '🎯';
    if (numScore >= 70) return '🎉';
    if (numScore >= 50) return '👍';
    return '📝';
  };

  if (loading && assignments.length === 0) {
    return (
      <div className="submit-scores-loading">
        <div className="loading-spinner"></div>
        <p>Loading your assignments...</p>
      </div>
    );
  }

  return (
    <div className="submit-scores">
      <div className="page-header">
        <h1 className="page-title">Submit Scores</h1>
        <p className="page-subtitle">Enter scores for students in your assigned classes</p>
      </div>

      <div className="submit-scores-card">
        <form onSubmit={handleSubmit}>
          {/* Success Message */}
          {successMessage && (
            <div className="form-success">
              <div className="success-icon">✓</div>
              <div className="success-message">{successMessage}</div>
            </div>
          )}

          {/* Error Message */}
          {error && (
            <div className="form-error">
              <div className="error-icon">⚠</div>
              <div className="error-content">
                <div className="error-message">{error}</div>
                {assignments.length === 0 && (
                  <div className="error-action">
                    <button
                      type="button"
                      className="btn btn-secondary btn-sm"
                      onClick={fetchAssignments}
                      disabled={!academicYear || !selectedTermId || submitting}
                    >
                      Retry Loading Assignments
                    </button>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Academic Year and Term Selection */}
          <div className="filters-row">
            <div className="form-field">
              <label htmlFor="academicYear" className="form-label">
                <span className="label-text">Academic Year</span>
                <span className="required-asterisk">*</span>
              </label>
              <select
                id="academicYear"
                value={academicYear}
                onChange={(e) => setAcademicYear(e.target.value)}
                disabled={submitting}
                className="styled-select"
              >
                <option value="2024">2024</option>
                <option value="2025">2025</option>
                <option value="2026">2026</option>
              </select>
            </div>

            <div className="form-field">
              <label htmlFor="term" className="form-label">
                <span className="label-text">Term</span>
                <span className="required-asterisk">*</span>
              </label>
              <select
                id="term"
                value={selectedTermId}
                onChange={(e) => setSelectedTermId(e.target.value)}
                disabled={submitting || availableTerms.length === 0}
                className="styled-select"
              >
                {availableTerms.map((term) => (
                  <option key={term.term_id} value={term.term_id}>
                    {term.term_name} {term.academic_year ? `(${term.academic_year})` : ''}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Assignment Selector */}
          {academicYear && selectedTermId && (
            <div className="form-field">
              <label htmlFor="assignment" className="form-label">
                <span className="label-text">Select Assignment</span>
                <span className="required-asterisk">*</span>
              </label>
              <select
                id="assignment"
                value={selectedAssignment}
                onChange={(e) => setSelectedAssignment(e.target.value)}
                disabled={submitting || loading || assignments.length === 0}
                className="styled-select"
              >
                <option value="">-- Choose an assignment --</option>
                {assignments.map((assignment) => (
                  <option key={assignment.assignment_id} value={assignment.assignment_id}>
                    {getAssignmentDisplayName(assignment)}
                  </option>
                ))}
              </select>
            </div>
          )}

          {/* Assignment Info */}
          {selectedAssignment && assignments.length > 0 && (
            <div className="assignment-info">
              <h3 className="assignment-title">Assignment Details</h3>
              <div className="assignment-details">
                {assignments
                  .filter((a) => a.assignment_id == selectedAssignment)
                  .map((assignment) => (
                    <div key={assignment.assignment_id} className="detail-flex-container">
                      <div className="detail-item">
                        <span className="detail-label">Subject:</span>
                        <span className="detail-value">{assignment.assignment_subject?.subject_name || 'N/A'}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Class:</span>
                        <span className="detail-value">{assignment.assignment_class?.class_name || 'N/A'}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Stream:</span>
                        <span className="detail-value">{assignment.assignment_stream?.stream_name || 'None'}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Term:</span>
                        <span className="detail-value">{assignment.assignment_term?.term_name || 'N/A'}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Academic Year:</span>
                        <span className="detail-value">{assignment.academic_year || 'N/A'}</span>
                      </div>
                    </div>
                  ))}
              </div>
            </div>
          )}

          {/* Students Table */}
          {selectedAssignment && students.length > 0 ? (
            <div className="students-section">
              <div className="section-header">
                <h3 className="section-title">Students</h3>
                <div className="student-count-badge">{students.length} students</div>
              </div>
              
              <div className="table-container">
                <table className="students-table">
                  <thead>
                    <tr>
                      <th className="table-header">#</th>
                      <th className="table-header">Admission No</th>
                      <th className="table-header">Student Name</th>
                      <th className="table-header score-header">Score (0-100)</th>
                      <th className="table-header status-header">Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {students.map((student, index) => (
                      <tr 
                        key={student.student_id} 
                        className={`table-row ${index % 2 === 0 ? 'even' : 'odd'}`}
                      >
                        <td className="table-cell serial-number" data-label="#">{index + 1}</td>
                        <td className="table-cell admission-number" data-label="Admission No">
                          {student.admission_number}
                        </td>
                        <td className="table-cell student-name" data-label="Student Name">
                          {student.fullname}
                        </td>
                        <td className="table-cell score-cell" data-label="Score">
                          <div className="score-input-wrapper">
                            <input
                              type="number"
                              min="0"
                              max="100"
                              step="0.1"
                              value={scores[student.student_id] || ''}
                              onChange={(e) => handleScoreChange(student.student_id, e.target.value)}
                              placeholder="0-100"
                              disabled={submitting}
                              className="score-input-table"
                            />
                          </div>
                        </td>
                        <td className="table-cell status-cell" data-label="Status">
                          {scores[student.student_id] ? (
                            <div className="score-status">
                              <span className="score-emoji">
                                {getScoreEmoji(scores[student.student_id])}
                              </span>
                              <span className="score-value">
                                {parseFloat(scores[student.student_id])}%
                              </span>
                            </div>
                          ) : (
                            <span className="no-score">Not entered</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Quick Stats */}
              <div className="score-stats">
                <div className="stat-item">
                  <span className="stat-label">Total Students:</span>
                  <span className="stat-value">{students.length}</span>
                </div>
                <div className="stat-item">
                  <span className="stat-label">Scores Entered:</span>
                  <span className="stat-value">
                    {Object.values(scores).filter(score => score !== '').length}
                  </span>
                </div>
                <div className="stat-item">
                  <span className="stat-label">Pending:</span>
                  <span className="stat-value">
                    {students.length - Object.values(scores).filter(score => score !== '').length}
                  </span>
                </div>
              </div>
            </div>
          ) : selectedAssignment ? (
            <div className="no-students">
              <div className="empty-state">
                <div className="empty-icon">👥</div>
                <h3>No Students Found</h3>
                <p>No students are currently enrolled in this assignment.</p>
              </div>
            </div>
          ) : academicYear && selectedTermId && assignments.length === 0 ? (
            <div className="no-assignments">
              <div className="empty-state">
                <div className="empty-icon">📚</div>
                <h3>No Teaching Assignments</h3>
                <p>
                  You don't have any teaching assignments for {academicYear} - {getTermName(selectedTermId)}.
                </p>
                <p className="empty-subtext">Please contact the administration to get assigned to classes.</p>
              </div>
            </div>
          ) : (
            <div className="no-selection">
              <div className="empty-state">
                <div className="empty-icon">🔍</div>
                <h3>Select Academic Details</h3>
                <p>Please select both academic year and term to view your assignments.</p>
              </div>
            </div>
          )}

          <div className="submit-actions">
            <button
              type="button"
              className="btn btn-secondary back-button"
              onClick={() => navigate('/teacher/assigned-subjects')}
              disabled={submitting}
            >
              ← Back to Assignments
            </button>

            {selectedAssignment && students.length > 0 && (
              <button
                type="submit"
                className="btn btn-primary submit-button"
                disabled={
                  submitting ||
                  !selectedAssignment ||
                  !selectedTermId ||
                  Object.values(scores).filter((s) => s !== '').length === 0 ||
                  isNaN(parseInt(selectedAssignment)) ||
                  isNaN(parseInt(selectedTermId))
                }
              >
                {submitting ? (
                  <>
                    <span className="button-spinner"></span>
                    Submitting...
                  </>
                ) : (
                  `📊 Submit ${Object.values(scores).filter(s => s !== '').length} Scores`
                )}
              </button>
            )}
          </div>
        </form>
      </div>
    </div>
  );
};

export default SubmitScores;