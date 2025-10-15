import React, { useState, useEffect } from 'react';
import { get, post, put, del } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Button from '../../components/ui/Button';
import Modal from '../../components/ui/Modal';
import Select from '../../components/ui/Select';
import Tabs from '../../components/ui/Tabs';
import Badge from '../../components/ui/Badge';
import './ManageSubjects.css';

const ManageSubjects = () => {
  const [formData, setFormData] = useState({
    subject_name: '',
    subject_code: '',
    subject_type: 'core'
  });
  
  const [editFormData, setEditFormData] = useState({
    subject_name: '',
    subject_code: '',
    subject_type: 'core',
    is_active: true
  });
  
  const [formErrors, setFormErrors] = useState({});
  const [isSubmitting, setIsSubmitting] = useState(false);

  const [subjects, setSubjects] = useState([]);
  const [teachers, setTeachers] = useState([]);
  const [classes, setClasses] = useState([]);
  const [streams, setStreams] = useState([]);
  const [terms, setTerms] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tableError, setTableError] = useState('');

  const [showAddModal, setShowAddModal] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const [deleteSubject, setDeleteSubject] = useState(null);
  const [modalLoading, setModalLoading] = useState(false);
  const [viewSubject, setViewSubject] = useState(null);
  const [assignTeacherModal, setAssignTeacherModal] = useState(null);
  const [assignmentForm, setAssignmentForm] = useState({
    teacher_id: '',
    class_id: '',
    stream_id: '',
    term_id: '',
    academic_year: new Date().getFullYear().toString()
  });
  
  const [activeTab, setActiveTab] = useState('all');
  const [stats, setStats] = useState(null);
  const [expandedSubjects, setExpandedSubjects] = useState(new Set());

  // Toggle subject details visibility
  const toggleSubjectDetails = (subjectId) => {
    const newExpanded = new Set(expandedSubjects);
    if (newExpanded.has(subjectId)) {
      newExpanded.delete(subjectId);
    } else {
      newExpanded.add(subjectId);
    }
    setExpandedSubjects(newExpanded);
  };

  // Fetch classes with proper error handling
  const fetchClasses = async () => {
    try {
      const response = await get('/classes?is_active=true');
      
      if (response.success) {
        const validClasses = (response.data || []).map(classItem => ({
          class_id: classItem.class_id,
          class_name: classItem.class_name || 'Unnamed Class',
          class_level: classItem.class_level || 'N/A',
          is_active: classItem.is_active !== false
        }));
        setClasses(validClasses);
      } else {
        console.error('Failed to fetch classes:', response.message);
        setClasses([]);
      }
    } catch (err) {
      console.error('Error fetching classes:', err);
      setClasses([]);
    }
  };

  const fetchData = async () => {
    try {
      setLoading(true);
      const [subjectsRes, teachersRes, streamsRes, termsRes] = await Promise.all([
        get('/subjects'),
        get('/teachers?is_active=true'),
        get('/streams?is_active=true'),
        get('/terms?is_active=true')
      ]);

      await fetchClasses();

      setSubjects(subjectsRes?.data || []);
      setTeachers(teachersRes?.data || []);
      setStreams(streamsRes?.data || []);
      setTerms(termsRes?.data || []);
      setTableError('');
    } catch (err) {
      console.error('[fetchData] Failed to fetch data:', err);
      setTableError('Failed to load data. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const fetchSubjectStats = async (subjectId) => {
    try {
      const response = await get(`/subjects/${subjectId}/stats`);
      setStats(response.data);
    } catch (err) {
      console.error('Error fetching subject stats:', err);
    }
  };

  const handleChange = (field, value) => {
    setFormData(prev => ({ ...prev, [field]: value }));
    if (formErrors[field]) setFormErrors(prev => ({ ...prev, [field]: '' }));
  };

  const validateForm = () => {
    const errors = {};
    if (!formData.subject_name.trim()) {
      errors.subject_name = 'Subject name is required';
    }
    if (!formData.subject_code.trim()) {
      errors.subject_code = 'Subject code is required';
    }
    setFormErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleAddSubmit = async (e) => {
    e.preventDefault();
    if (isSubmitting) return;
    if (!validateForm()) return;

    setIsSubmitting(true);
    try {
      const response = await post('/subjects', formData);
      if (response.success) {
        setFormData({ subject_name: '', subject_code: '', subject_type: 'core' });
        setFormErrors({});
        setShowAddModal(false);
        fetchData();
        alert('Subject created successfully!');
      } else {
        setFormErrors({ submit: response.message || 'Failed to add subject' });
      }
    } catch (err) {
      const errorMessage = err.response?.data?.message || 'Failed to add subject. Please try again.';
      setFormErrors({ submit: errorMessage });
      console.error('[handleAddSubmit] Error:', err);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleEditSubmit = async (e) => {
    e.preventDefault();
    if (isSubmitting) return;

    setIsSubmitting(true);
    try {
      const response = await put(`/subjects/${viewSubject.subject_id}`, editFormData);
      if (response.success) {
        setShowEditModal(false);
        fetchData();
        alert('Subject updated successfully!');
      } else {
        alert(response.message || 'Failed to update subject');
      }
    } catch (err) {
      const errorMessage = err.response?.data?.message || 'Failed to update subject.';
      alert(errorMessage);
    } finally {
      setIsSubmitting(false);
    }
  };

  const isDuplicateAssignment = () => {
    if (
      !assignmentForm.teacher_id ||
      !assignmentForm.class_id ||
      !assignmentForm.stream_id ||
      !assignmentForm.term_id ||
      !assignTeacherModal
    ) {
      return false;
    }

    const subjectData = subjects.find((s) => s.subject_id === assignTeacherModal.subject_id);
    return subjectData?.teachers?.some(
      (teacher) =>
        teacher.teacher_id === parseInt(assignmentForm.teacher_id) &&
        teacher.classes?.some(
          (cls) =>
            cls.class_id === parseInt(assignmentForm.class_id) &&
            cls.stream_id === parseInt(assignmentForm.stream_id) &&
            cls.term_id === parseInt(assignmentForm.term_id)
        )
    );
  };

  const handleCreateAssignment = async (subject) => {
    if (isSubmitting) return;

    if (
      !assignmentForm.teacher_id ||
      !assignmentForm.class_id ||
      !assignmentForm.stream_id ||
      !assignmentForm.term_id
    ) {
      alert('Please fill all required fields including stream');
      return;
    }

    const assignmentData = {
      teacher_id: parseInt(assignmentForm.teacher_id),
      subject_id: subject.subject_id,
      class_id: parseInt(assignmentForm.class_id),
      stream_id: parseInt(assignmentForm.stream_id),
      term_id: parseInt(assignmentForm.term_id),
      academic_year: assignmentForm.academic_year,
      is_active: true
    };

    // Check for duplicate assignment
    const subjectData = subjects.find((s) => s.subject_id === subject.subject_id);
    const existingAssignment = subjectData?.teachers?.some(
      (teacher) =>
        teacher.teacher_id === assignmentData.teacher_id &&
        teacher.classes?.some(
          (cls) =>
            cls.class_id === assignmentData.class_id &&
            cls.stream_id === assignmentData.stream_id &&
            cls.term_id === assignmentData.term_id
        )
    );

    if (existingAssignment) {
      alert('This teacher is already assigned to this subject, class, stream, and term.');
      return;
    }

    setIsSubmitting(true);
    try {
      console.log('Sending assignment data:', assignmentData);
      const response = await post('/assignments', assignmentData);
      if (response.success) {
        setAssignTeacherModal(null);
        setAssignmentForm({
          teacher_id: '',
          class_id: '',
          stream_id: '',
          term_id: '',
          academic_year: new Date().getFullYear().toString()
        });
        fetchData();
        alert('Teacher assigned successfully!');
      } else {
        alert(response.message || 'Failed to assign teacher');
      }
    } catch (err) {
      const errorMessage =
        err.response?.data?.message === 'Assignment already exists'
          ? 'This teacher is already assigned to this subject, class, stream, and term.'
          : err.response?.data?.message || 'Failed to assign teacher';
      alert(errorMessage);
      console.error('Assignment error:', err.response?.data);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleRemoveAssignment = async (subjectId, assignmentId) => {
    if (!window.confirm('Are you sure you want to remove this assignment?')) return;

    try {
      const response = await del(`/assignments/${assignmentId}`);
      if (response.success) {
        fetchData();
        alert('Assignment removed successfully!');
      } else {
        alert(response.message || 'Failed to remove assignment');
      }
    } catch (err) {
      alert(err.response?.data?.message || 'Failed to remove assignment');
    }
  };

  const handleDelete = async () => {
    if (!deleteSubject) return;
    setModalLoading(true);
    try {
      const response = await del(`/subjects/${deleteSubject.subject_id}`);
      if (response.success) {
        setDeleteSubject(null);
        fetchData();
        alert('Subject deactivated successfully!');
      } else {
        alert(response.message || 'Failed to delete subject');
      }
    } catch (err) {
      const errorMessage = err.response?.data?.message || 'Failed to delete subject';
      alert(errorMessage);
    } finally {
      setModalLoading(false);
    }
  };

  const handleReactivate = async (subject) => {
    try {
      const response = await put(`/subjects/${subject.subject_id}/reactivate`);
      if (response.success) {
        fetchData();
        alert('Subject reactivated successfully!');
      } else {
        alert(response.message || 'Failed to reactivate subject');
      }
    } catch (err) {
      alert(err.response?.data?.message || 'Failed to reactivate subject');
    }
  };

  const getAssignmentOverview = (subject) => {
    const teachers = subject.teachers || [];
    const assignmentCount = teachers.reduce((count, teacher) => 
      count + (teacher.classes || []).length, 0
    );

    if (assignmentCount === 0) {
      return <div className="no-assignments-overview">No teachers assigned</div>;
    }

    const teacherNames = teachers.map(teacher => 
      teacher.user_account?.fullname || teacher.fullname
    ).join(', ');

    return (
      <div className="assignment-overview">
        <div className="overview-stats">
          <Badge variant="primary" className="assignment-count">
            {assignmentCount} assignment{assignmentCount !== 1 ? 's' : ''}
          </Badge>
          <span className="teacher-names">{teacherNames}</span>
        </div>
      </div>
    );
  };

  const getTeacherAssignments = (subject) => {
    const teachers = subject.teachers || [];
    
    if (teachers.length === 0) {
      return <div className="no-assignments">No teachers assigned</div>;
    }
    
    return teachers.map(teacher => (
      <div key={teacher.teacher_id} className="teacher-assignment">
        <div className="teacher-info">
          <strong>{teacher.user_account?.fullname || teacher.fullname}</strong> ({teacher.teacher_code})
        </div>
        <div className="assignment-details">
          {teacher.classes && teacher.classes.map((cls, index) => (
            <Badge key={index} variant="outline" className="assignment-badge">
              {cls.class_name} {cls.stream_name ? `- ${cls.stream_name}` : ''} ({cls.term_name})
              <button 
                type="button"
                className="remove-assignment-btn"
                onClick={() => handleRemoveAssignment(subject.subject_id, cls.assignment_id)}
                title="Remove assignment"
              >
                ×
              </button>
            </Badge>
          ))}
        </div>
      </div>
    ));
  };

  const openEditModal = (subject) => {
    setViewSubject(subject);
    setEditFormData({
      subject_name: subject.subject_name,
      subject_code: subject.subject_code,
      subject_type: subject.subject_type,
      is_active: subject.is_active
    });
    setShowEditModal(true);
  };

  const openAssignTeacherModal = (subject) => {
    setAssignTeacherModal(subject);
    setAssignmentForm({
      teacher_id: '',
      class_id: '',
      stream_id: '',
      term_id: '',
      academic_year: new Date().getFullYear().toString()
    });
  };

  const openViewModal = async (subject) => {
    setViewSubject(subject);
    await fetchSubjectStats(subject.subject_id);
  };

  const filteredSubjects = subjects.filter(subject => {
    switch (activeTab) {
      case 'active':
        return subject.is_active !== false;
      case 'inactive':
        return subject.is_active === false;
      default:
        return true;
    }
  });

  const classStreams = assignmentForm.class_id 
    ? streams.filter(stream => stream.class_id === parseInt(assignmentForm.class_id))
    : [];

  // Loading state
  if (loading) {
    return (
      <div className="manage-subjects">
        <div className="page-header">
          <h1 className="page-title">Manage Subjects</h1>
        </div>
        <Card>
          <div className="loading-state">Loading subjects...</div>
        </Card>
      </div>
    );
  }

  return (
    <div className="manage-subjects">
      <div className="page-header">
        <h1 className="page-title">Manage Subjects</h1>
        <Button onClick={() => setShowAddModal(true)} className="add-subject-btn">
          + Add Subject
        </Button>
      </div>

      <Card>
        <Tabs
          tabs={[
            { id: 'all', label: 'All Subjects' },
            { id: 'active', label: 'Active' },
            { id: 'inactive', label: 'Inactive' }
          ]}
          activeTab={activeTab}
          onChange={setActiveTab}
        />
        
        {tableError && (
          <div className="error-message">
            {tableError}
            <Button variant="secondary" size="sm" onClick={fetchData} style={{ marginLeft: '10px' }}>
              Retry
            </Button>
          </div>
        )}
        
        {/* Subjects Table */}
        <div className="subjects-table-container">
          {filteredSubjects.length === 0 ? (
            <div className="empty-state">
              <p>No subjects found. Add your first subject!</p>
            </div>
          ) : (
            <>
              {/* Desktop Table */}
              <table className="subjects-table desktop-table">
                <thead>
                  <tr>
                    <th>Subject Name</th>
                    <th>Assignments</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSubjects.map((subject) => (
                    <React.Fragment key={subject.subject_id}>
                      <tr className="subject-row">
                        <td>
                          <div className="subject-info-cell">
                            <div className="subject-main-info">
                              <span className="subject-title">{subject.subject_name}</span>
                              <span className="subject-code">{subject.subject_code}</span>
                            </div>
                            <div className="subject-badges">
                              <Badge variant={subject.subject_type === 'core' ? 'primary' : 'secondary'}>
                                {subject.subject_type}
                              </Badge>
                              {!subject.is_active && <Badge variant="danger">Inactive</Badge>}
                            </div>
                          </div>
                        </td>
                        <td>
                          {getAssignmentOverview(subject)}
                        </td>
                        <td>
                          <div className="action-buttons">
                            <Button
                              size="sm"
                              variant="outline"
                              onClick={() => toggleSubjectDetails(subject.subject_id)}
                              className="details-btn"
                            >
                              {expandedSubjects.has(subject.subject_id) ? '▲ Hide' : '▼ Details'}
                            </Button>
                            <Button
                              size="sm"
                              variant="primary"
                              onClick={() => openAssignTeacherModal(subject)}
                              disabled={!subject.is_active}
                            >
                              Assign Teacher
                            </Button>
                            <Button
                              size="sm"
                              variant="secondary"
                              onClick={() => openEditModal(subject)}
                            >
                              Edit
                            </Button>
                            {subject.is_active ? (
                              <Button
                                size="sm"
                                variant="danger"
                                onClick={() => setDeleteSubject(subject)}
                              >
                                Deactivate
                              </Button>
                            ) : (
                              <Button
                                size="sm"
                                variant="success"
                                onClick={() => handleReactivate(subject)}
                              >
                                Reactivate
                              </Button>
                            )}
                          </div>
                        </td>
                      </tr>

                      {/* Expanded Details Row */}
                      {expandedSubjects.has(subject.subject_id) && (
                        <tr className="expanded-details-row">
                          <td colSpan="3">
                            <div className="subject-details-expanded">
                              <div className="details-header">
                                <h4>Teacher Assignments for {subject.subject_name}</h4>
                                <Button 
                                  size="sm" 
                                  variant="outline" 
                                  onClick={() => toggleSubjectDetails(subject.subject_id)}
                                >
                                  ▲ Hide Details
                                </Button>
                              </div>
                              <div className="assignments-container">
                                {getTeacherAssignments(subject)}
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>

              {/* Mobile Cards */}
              <div className="mobile-subjects-list">
                {filteredSubjects.map((subject) => (
                  <div key={subject.subject_id} className="subject-card">
                    <div className="subject-card-header">
                      <div className="subject-card-main">
                        <div className="subject-card-name">{subject.subject_name}</div>
                        <div className="subject-card-code">{subject.subject_code}</div>
                      </div>
                      <div className="subject-card-status">
                        <Badge variant={subject.subject_type === 'core' ? 'primary' : 'secondary'}>
                          {subject.subject_type}
                        </Badge>
                        {!subject.is_active && <Badge variant="danger">Inactive</Badge>}
                      </div>
                    </div>
                    
                    <div className="subject-card-details">
                      <div className="assignment-overview">
                        {getAssignmentOverview(subject)}
                      </div>
                    </div>

                    <div className="subject-card-actions">
                      <div className="action-buttons">
                        <Button
                          size="sm"
                          variant="outline"
                          onClick={() => toggleSubjectDetails(subject.subject_id)}
                          fullWidth
                        >
                          {expandedSubjects.has(subject.subject_id) ? '▲ Hide' : '▼ Details'}
                        </Button>
                        <Button
                          size="sm"
                          variant="primary"
                          onClick={() => openAssignTeacherModal(subject)}
                          disabled={!subject.is_active}
                          fullWidth
                        >
                          Assign Teacher
                        </Button>
                        <Button
                          size="sm"
                          variant="secondary"
                          onClick={() => openEditModal(subject)}
                          fullWidth
                        >
                          Edit
                        </Button>
                        {subject.is_active ? (
                          <Button
                            size="sm"
                            variant="danger"
                            onClick={() => setDeleteSubject(subject)}
                            fullWidth
                          >
                            Deactivate
                          </Button>
                        ) : (
                          <Button
                            size="sm"
                            variant="success"
                            onClick={() => handleReactivate(subject)}
                            fullWidth
                          >
                            Reactivate
                          </Button>
                        )}
                      </div>
                    </div>

                    {/* Expanded Details for Mobile */}
                    {expandedSubjects.has(subject.subject_id) && (
                      <div className="subject-details-expanded-mobile">
                        <div className="details-header">
                          <h4>Teacher Assignments</h4>
                        </div>
                        <div className="assignments-container">
                          {getTeacherAssignments(subject)}
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </Card>

      {/* Add Subject Modal */}
      <Modal isOpen={showAddModal} onClose={() => setShowAddModal(false)} title="Add New Subject" size="md">
        <form onSubmit={handleAddSubmit} className="subject-form">
          {formErrors.submit && <div className="form-error">{formErrors.submit}</div>}

          <div className="form-fields">
            <Input 
              label="Subject Name *" 
              value={formData.subject_name} 
              onChange={(e) => handleChange('subject_name', e.target.value)} 
              error={formErrors.subject_name} 
              required 
            />
            
            <Input 
              label="Subject Code *" 
              value={formData.subject_code} 
              onChange={(e) => handleChange('subject_code', e.target.value)} 
              error={formErrors.subject_code} 
              required 
            />

            <Select 
              label="Subject Type *" 
              value={formData.subject_type} 
              onChange={(e) => handleChange('subject_type', e.target.value)}
            >
              <option value="core">Core</option>
              <option value="elective">Elective</option>
              <option value="optional">Optional</option>
            </Select>
          </div>

          <div className="form-actions">
            <Button type="button" variant="secondary" onClick={() => setShowAddModal(false)} disabled={isSubmitting}>
              Cancel
            </Button>
            <Button type="submit" loading={isSubmitting}>
              {isSubmitting ? 'Creating...' : 'Create Subject'}
            </Button>
          </div>
        </form>
      </Modal>

      {/* Edit Subject Modal */}
      <Modal isOpen={showEditModal} onClose={() => setShowEditModal(false)} title="Edit Subject" size="md">
        <form onSubmit={handleEditSubmit} className="subject-form">
          <div className="form-fields">
            <Input 
              label="Subject Name *" 
              value={editFormData.subject_name} 
              onChange={(e) => setEditFormData(prev => ({ ...prev, subject_name: e.target.value }))} 
              required 
            />
            
            <Input 
              label="Subject Code *" 
              value={editFormData.subject_code} 
              onChange={(e) => setEditFormData(prev => ({ ...prev, subject_code: e.target.value }))} 
              required 
            />

            <Select 
              label="Subject Type *" 
              value={editFormData.subject_type} 
              onChange={(e) => setEditFormData(prev => ({ ...prev, subject_type: e.target.value }))}
            >
              <option value="core">Core</option>
              <option value="elective">Elective</option>
              <option value="optional">Optional</option>
            </Select>

            <Select 
              label="Status" 
              value={editFormData.is_active} 
              onChange={(e) => setEditFormData(prev => ({ ...prev, is_active: e.target.value === 'true' }))}
            >
              <option value={true}>Active</option>
              <option value={false}>Inactive</option>
            </Select>
          </div>

          <div className="form-actions">
            <Button type="button" variant="secondary" onClick={() => setShowEditModal(false)} disabled={isSubmitting}>
              Cancel
            </Button>
            <Button type="submit" loading={isSubmitting}>
              {isSubmitting ? 'Updating...' : 'Update Subject'}
            </Button>
          </div>
        </form>
      </Modal>

      {/* Assign Teacher Modal */}
      <Modal 
        isOpen={!!assignTeacherModal} 
        onClose={() => setAssignTeacherModal(null)} 
        title={`Assign Teacher - ${assignTeacherModal?.subject_name}`}
        size="md"
      >
        <div className="assign-teacher-form">
          <Select 
            label="Teacher *" 
            value={assignmentForm.teacher_id} 
            onChange={(e) => setAssignmentForm(prev => ({ ...prev, teacher_id: e.target.value }))}
          >
            <option value="">-- Select Teacher --</option>
            {teachers.map(teacher => (
              <option key={teacher.teacher_id} value={teacher.teacher_id}>
                {teacher.user_account?.fullname || teacher.fullname} ({teacher.teacher_code})
              </option>
            ))}
          </Select>

          <Select 
            label="Class *" 
            value={assignmentForm.class_id} 
            onChange={(e) => setAssignmentForm(prev => ({ ...prev, class_id: e.target.value, stream_id: '' }))}
          >
            <option value="">-- Select Class --</option>
            {classes.map(classItem => (
              <option key={classItem.class_id} value={classItem.class_id}>
                {classItem.class_name} (Level: {classItem.class_level})
              </option>
            ))}
          </Select>

          <Select 
            label="Stream *" 
            value={assignmentForm.stream_id} 
            onChange={(e) => setAssignmentForm(prev => ({ ...prev, stream_id: e.target.value }))}
          >
            <option value="">-- Select Stream --</option>
            {classStreams.map(stream => (
              <option key={stream.stream_id} value={stream.stream_id}>
                {stream.stream_name}
              </option>
            ))}
          </Select>

          <Select 
            label="Term *" 
            value={assignmentForm.term_id} 
            onChange={(e) => setAssignmentForm(prev => ({ ...prev, term_id: e.target.value }))}
          >
            <option value="">-- Select Term --</option>
            {terms.map(term => (
              <option key={term.term_id} value={term.term_id}>
                {term.term_name} - {term.academic_year}
              </option>
            ))}
          </Select>

          <Input 
            label="Academic Year *" 
            type="number"
            value={assignmentForm.academic_year} 
            onChange={(e) => setAssignmentForm(prev => ({ ...prev, academic_year: e.target.value }))}
            required
          />

          <div className="form-actions">
            <Button 
              type="button" 
              variant="secondary" 
              onClick={() => setAssignTeacherModal(null)}
              disabled={isSubmitting}
            >
              Cancel
            </Button>
            <Button 
              type="button" 
              variant="primary"
              onClick={() => handleCreateAssignment(assignTeacherModal)}
              disabled={
                isSubmitting ||
                !assignmentForm.teacher_id ||
                !assignmentForm.class_id ||
                !assignmentForm.stream_id ||
                !assignmentForm.term_id ||
                isDuplicateAssignment()
              }
              loading={isSubmitting}
            >
              {isSubmitting ? 'Assigning...' : 'Assign Teacher'}
            </Button>
          </div>
        </div>
      </Modal>

      {/* View Subject Details Modal */}
      <Modal 
        isOpen={!!viewSubject && !showEditModal} 
        onClose={() => setViewSubject(null)} 
        title="Subject Details"
        size="lg"
      >
        {viewSubject && (
          <div className="subject-details">
            <div className="detail-section">
              <h3>Basic Information</h3>
              <div className="detail-grid">
                <div className="detail-item">
                  <strong>Subject Name:</strong> {viewSubject.subject_name}
                </div>
                <div className="detail-item">
                  <strong>Subject Code:</strong> {viewSubject.subject_code}
                </div>
                <div className="detail-item">
                  <strong>Subject Type:</strong> 
                  <Badge variant={viewSubject.subject_type === 'core' ? 'primary' : 'secondary'}>
                    {viewSubject.subject_type}
                  </Badge>
                </div>
                <div className="detail-item">
                  <strong>Status:</strong> 
                  <Badge variant={viewSubject.is_active ? 'success' : 'danger'}>
                    {viewSubject.is_active ? 'Active' : 'Inactive'}
                  </Badge>
                </div>
              </div>
            </div>

            {stats && (
              <div className="detail-section">
                <h3>Statistics</h3>
                <div className="stats-grid">
                  <div className="stat-item">
                    <strong>Total Assignments:</strong> {stats.assignment_count || 0}
                  </div>
                  <div className="stat-item">
                    <strong>Total Scores:</strong> {stats.score_count || 0}
                  </div>
                  <div className="stat-item">
                    <strong>Average Score:</strong> {stats.average_score ? Math.round(stats.average_score) : 'N/A'}
                  </div>
                </div>
              </div>
            )}

            <div className="detail-section">
              <h3>Teacher Assignments</h3>
              <div className="assignments-list">
                {getTeacherAssignments(viewSubject)}
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* Delete Confirmation Modal */}
      <Modal isOpen={!!deleteSubject} onClose={() => setDeleteSubject(null)} title="Confirm Deactivation" size="sm">
        <p>Are you sure you want to deactivate <strong>{deleteSubject?.subject_name}</strong>?</p>
        <p className="text-muted">
          This will make the subject unavailable for new assignments but will preserve existing data.
        </p>
        <div className="modal-actions">
          <Button variant="secondary" onClick={() => setDeleteSubject(null)} disabled={modalLoading}>
            Cancel
          </Button>
          <Button variant="danger" onClick={handleDelete} loading={modalLoading}>
            Deactivate
          </Button>
        </div>
      </Modal>
    </div>
  );
};

export default ManageSubjects;