import React, { useState, useEffect, useCallback } from 'react';
import { get, post, put, del } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Button from '../../components/ui/Button';
import Modal from '../../components/ui/Modal';
import Select from '../../components/ui/Select';
import Tabs from '../../components/ui/Tabs';
import Badge from '../../components/ui/Badge';
import './ManageTeachers.css';

const FORM_INITIAL_STATE = {
  fullname: '',        
  email: '',
  phone_number: '',
  password: '',
  confirmPassword: '',
  teacher_code: '',
  specialization: ''
};

const useTeachers = () => {
  const [teachers, setTeachers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  const fetchTeachers = useCallback(async (is_active = true) => {
    try {
      setLoading(true);
      setError('');
      const response = await get(`/teachers?is_active=${is_active}`);
      setTeachers(response?.data || []);
    } catch (err) {
      console.error('❌ Failed to fetch teachers:', err);
      setError(err.response?.data?.message || 'Failed to load teachers');
      setTeachers([]);
    } finally {
      setLoading(false);
    }
  }, []);

  const addTeacher = useCallback(async (teacherData) => {
    try {
      const response = await post('/teachers', teacherData);
      return { success: true, data: response };
    } catch (error) {
      return { 
        success: false, 
        error: error.response?.data?.message || 'Failed to add teacher' 
      };
    }
  }, []);

  const updateTeacher = useCallback(async (teacherId, teacherData) => {
    try {
      const response = await put(`/teachers/${teacherId}`, teacherData);
      return { success: true, data: response };
    } catch (error) {
      return { 
        success: false, 
        error: error.response?.data?.message || 'Failed to update teacher' 
      };
    }
  }, []);

  const deleteTeacher = useCallback(async (teacherId) => {
    try {
      const response = await del(`/teachers/${teacherId}`);
      return { success: true, data: response };
    } catch (error) {
      return { 
        success: false, 
        error: error.response?.data?.message || 'Failed to delete teacher' 
      };
    }
  }, []);

  const reactivateTeacher = useCallback(async (teacherId) => {
    try {
      const response = await put(`/teachers/${teacherId}/reactivate`);
      return { success: true, data: response };
    } catch (error) {
      return { 
        success: false, 
        error: error.response?.data?.message || 'Failed to reactivate teacher' 
      };
    }
  }, []);

  return {
    teachers,
    loading,
    error,
    setError,
    fetchTeachers,
    addTeacher,
    updateTeacher,
    deleteTeacher,
    reactivateTeacher
  };
};

const useTeacherDetails = () => {
  const [teacherDetails, setTeacherDetails] = useState(null);
  const [assignments, setAssignments] = useState([]);
  const [students, setStudents] = useState([]);
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(false);

  const fetchTeacherDetails = useCallback(async (teacherId) => {
    try {
      setLoading(true);
      const response = await get(`/teachers/${teacherId}`);
      setTeacherDetails(response.data);
    } catch (err) {
      console.error('Error fetching teacher details:', err);
    } finally {
      setLoading(false);
    }
  }, []);

  const fetchTeacherAssignments = useCallback(async (teacherId, filters = {}) => {
    try {
      const queryParams = new URLSearchParams(filters).toString();
      const response = await get(`/teachers/${teacherId}/assignments?${queryParams}`);
      setAssignments(response.data || []);
    } catch (err) {
      console.error('Error fetching teacher assignments:', err);
      setAssignments([]);
    }
  }, []);

  const fetchTeacherStudents = useCallback(async (teacherId) => {
    try {
      const response = await get(`/teachers/${teacherId}/students`);
      setStudents(response.data || []);
    } catch (err) {
      console.error('Error fetching teacher students:', err);
      setStudents([]);
    }
  }, []);

  const fetchTeacherStats = useCallback(async (teacherId) => {
    try {
      const response = await get(`/teachers/${teacherId}/stats`);
      setStats(response.data);
    } catch (err) {
      console.error('Error fetching teacher stats:', err);
      setStats(null);
    }
  }, []);

  return {
    teacherDetails,
    assignments,
    students,
    stats,
    loading,
    fetchTeacherDetails,
    fetchTeacherAssignments,
    fetchTeacherStudents,
    fetchTeacherStats
  };
};

const ManageTeachers = () => {
  // State management
  const [formData, setFormData] = useState(FORM_INITIAL_STATE);
  const [formErrors, setFormErrors] = useState({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [modalLoading, setModalLoading] = useState(false);
  
  // Modal states
  const [showAddModal, setShowAddModal] = useState(false);
  const [editTeacher, setEditTeacher] = useState(null);
  const [deleteTeacher, setDeleteTeacher] = useState(null);
  const [viewTeacher, setViewTeacher] = useState(null);
  const [reactivateTeacher, setReactivateTeacher] = useState(null);
  
  // View states
  const [activeTab, setActiveTab] = useState('all');
  const [assignmentFilters, setAssignmentFilters] = useState({
    term_id: '',
    academic_year: new Date().getFullYear().toString()
  });

  // Custom hooks
  const {
    teachers,
    loading,
    error: tableError,
    fetchTeachers,
    addTeacher,
    updateTeacher,
    deleteTeacher: deleteTeacherApi,
    reactivateTeacher: reactivateTeacherApi
  } = useTeachers();

  const {
    teacherDetails,
    assignments,
    students,
    stats,
    fetchTeacherDetails,
    fetchTeacherAssignments,
    fetchTeacherStudents,
    fetchTeacherStats
  } = useTeacherDetails();

  // Fetch data
  useEffect(() => {
    fetchTeachers(activeTab === 'inactive' ? false : true);
  }, [fetchTeachers, activeTab]);

  // Form handlers
  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
    if (formErrors[name]) {
      setFormErrors(prev => ({ ...prev, [name]: '' }));
    }
  };

  const handleEditChange = (e) => {
    const { name, value } = e.target;
    setEditTeacher(prev => ({ ...prev, [name]: value }));
  };

  // Reset form function
  const resetForm = useCallback(() => {
    setFormData(FORM_INITIAL_STATE);
    setFormErrors({});
  }, []);

  const validateForm = (isEdit = false) => {
    const errors = {};

    if (!formData.fullname?.trim()) {
      errors.fullname = 'Full name is required';
    }

    if (!formData.email?.trim()) {
      errors.email = 'Email is required';
    } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(formData.email)) {
      errors.email = 'Please enter a valid email address';
    }

    if (!formData.teacher_code?.trim()) {
      errors.teacher_code = 'Teacher code is required';
    }

    if (!isEdit) {
      if (!formData.password) {
        errors.password = 'Password is required';
      } else if (formData.password.length < 6) {
        errors.password = 'Password must be at least 6 characters';
      }

      if (!formData.confirmPassword) {
        errors.confirmPassword = 'Please confirm your password';
      } else if (formData.password !== formData.confirmPassword) {
        errors.confirmPassword = 'Passwords do not match';
      }
    }

    setFormErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleAddSubmit = async (e) => {
    e.preventDefault();
    if (!validateForm()) return;

    setIsSubmitting(true);
    
    // Prepare data without employment_date
    const teacherData = {
      fullname: formData.fullname.trim(),
      email: formData.email.trim().toLowerCase(),
      password: formData.password,
      teacher_code: formData.teacher_code.trim(),
      phone_number: formData.phone_number?.trim() || null,
      specialization: formData.specialization?.trim() || null
    };

    const result = await addTeacher(teacherData);
    
    if (result.success) {
      resetForm();
      setShowAddModal(false);
      await fetchTeachers();
      alert('Teacher created successfully!');
    } else {
      setFormErrors({ submit: result.error });
    }
    
    setIsSubmitting(false);
  };

  const handleEditSubmit = async (e) => {
    e.preventDefault();
    if (!editTeacher) return;

    setModalLoading(true);
    
    // Prepare edit data without employment_date
    const teacherData = {
      fullname: editTeacher.fullname?.trim(),
      email: editTeacher.email?.trim().toLowerCase(),
      teacher_code: editTeacher.teacher_code?.trim(),
      phone_number: editTeacher.phone_number?.trim() || null,
      specialization: editTeacher.specialization?.trim() || null,
      is_active: editTeacher.is_active
    };

    const result = await updateTeacher(editTeacher.teacher_id, teacherData);
    
    if (result.success) {
      setEditTeacher(null);
      setFormErrors({});
      await fetchTeachers();
      alert('Teacher updated successfully!');
    } else {
      setFormErrors({ submit: result.error });
    }
    
    setModalLoading(false);
  };

  const handleDelete = async () => {
    if (!deleteTeacher) return;

    setModalLoading(true);
    const result = await deleteTeacherApi(deleteTeacher.teacher_id);
    
    if (result.success) {
      setDeleteTeacher(null);
      fetchTeachers();
      alert('Teacher deactivated successfully!');
    } else {
      setFormErrors({ submit: result.error });
    }
    setModalLoading(false);
  };

  const handleReactivate = async () => {
    if (!reactivateTeacher) return;

    setModalLoading(true);
    const result = await reactivateTeacherApi(reactivateTeacher.teacher_id);
    
    if (result.success) {
      setReactivateTeacher(null);
      fetchTeachers();
      alert('Teacher reactivated successfully!');
    } else {
      setFormErrors({ submit: result.error });
    }
    setModalLoading(false);
  };

  const handleViewTeacher = async (teacher) => {
    setViewTeacher(teacher);
    await Promise.all([
      fetchTeacherDetails(teacher.teacher_id),
      fetchTeacherAssignments(teacher.teacher_id),
      fetchTeacherStudents(teacher.teacher_id),
      fetchTeacherStats(teacher.teacher_id)
    ]);
  };

  const handleAssignmentFilterChange = (filter, value) => {
    setAssignmentFilters(prev => ({ ...prev, [filter]: value }));
    if (viewTeacher) {
      fetchTeacherAssignments(viewTeacher.teacher_id, { ...assignmentFilters, [filter]: value });
    }
  };

  // Helper functions
  const getTeacherDisplayName = (teacher) => {
    return teacher.user_account?.fullname || teacher.fullname || 'N/A';
  };

  const getTeacherEmail = (teacher) => {
    return teacher.user_account?.email || teacher.email || 'N/A';
  };

  const generateTeacherCode = () => {
    const newCode = `TCH${Date.now().toString().slice(-6)}`;
    setFormData(prev => ({ ...prev, teacher_code: newCode }));
  };

  const filteredTeachers = teachers.filter(teacher => {
    switch (activeTab) {
      case 'active':
        return teacher.is_active !== false;
      case 'inactive':
        return teacher.is_active === false;
      default:
        return true;
    }
  });

  // Loading state
  if (loading) {
    return (
      <div className="manage-teachers">
        <div className="page-header">
          <h1 className="page-title">Manage Teachers</h1>
        </div>
        <Card>
          <div className="loading-state">Loading teachers...</div>
        </Card>
      </div>
    );
  }

  return (
    <div className="manage-teachers">
      <div className="page-header">
        <h1 className="page-title">Manage Teachers</h1>
        <Button onClick={() => setShowAddModal(true)} className="add-teacher-btn">
          + Add Teacher
        </Button>
      </div>

      <Card>
        <Tabs
          tabs={[
            { id: 'all', label: 'All Teachers' },
            { id: 'active', label: 'Active' },
            { id: 'inactive', label: 'Inactive' }
          ]}
          activeTab={activeTab}
          onChange={setActiveTab}
        />
        
        {tableError && (
          <div className="error-message">
            {tableError}
            <Button variant="secondary" size="sm" onClick={() => fetchTeachers()} style={{ marginLeft: '10px' }}>
              Retry
            </Button>
          </div>
        )}
        
        {/* Teachers Table */}
        <div className="teachers-table-container">
          {filteredTeachers.length === 0 ? (
            <div className="empty-state">
              <p>No teachers found.</p>
            </div>
          ) : (
            <>
              {/* Desktop Table */}
              <table className="teachers-table desktop-table">
                <thead>
                  <tr>
                    <th>Teacher Information</th>
                    <th>Phone Number</th>
                    <th>Specialization</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredTeachers.map((teacher) => (
                    <tr key={teacher.teacher_id} className="teacher-row">
                      <td>
                        <div className="teacher-info-cell">
                          <div className="teacher-name">{getTeacherDisplayName(teacher)}</div>
                          <div className="teacher-details">
                            <span className="teacher-code">{teacher.teacher_code}</span>
                            <span className="teacher-email">{getTeacherEmail(teacher)}</span>
                          </div>
                          <div className="teacher-status">
                            {!teacher.is_active && <Badge variant="danger">Inactive</Badge>}
                            {teacher.specialization && (
                              <Badge variant="outline">{teacher.specialization}</Badge>
                            )}
                          </div>
                        </div>
                      </td>
                      <td>
                        {teacher.phone_number || 'Not provided'}
                      </td>
                      <td>
                        {teacher.specialization || 'Not specified'}
                      </td>
                      <td>
                        <div className="action-buttons">
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => handleViewTeacher(teacher)}
                          >
                            View
                          </Button>
                          <Button
                            size="sm"
                            variant="secondary"
                            onClick={() => setEditTeacher({
                              teacher_id: teacher.teacher_id,
                              fullname: getTeacherDisplayName(teacher),
                              email: getTeacherEmail(teacher),
                              teacher_code: teacher.teacher_code,
                              phone_number: teacher.phone_number || '',
                              specialization: teacher.specialization || '',
                              is_active: teacher.is_active
                            })}
                          >
                            Edit
                          </Button>
                          {teacher.is_active ? (
                            <Button
                              size="sm"
                              variant="danger"
                              onClick={() => setDeleteTeacher(teacher)}
                            >
                              Deactivate
                            </Button>
                          ) : (
                            <Button
                              size="sm"
                              variant="success"
                              onClick={() => setReactivateTeacher(teacher)}
                            >
                              Reactivate
                            </Button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Mobile Cards */}
              <div className="mobile-teachers-list">
                {filteredTeachers.map((teacher) => (
                  <div key={teacher.teacher_id} className="teacher-card">
                    <div className="teacher-card-header">
                      <div className="teacher-card-name">{getTeacherDisplayName(teacher)}</div>
                      <div className="teacher-card-status">
                        {!teacher.is_active && <Badge variant="danger">Inactive</Badge>}
                      </div>
                    </div>
                    
                    <div className="teacher-card-details">
                      <div className="detail-item">
                        <span className="detail-label">Teacher Code:</span>
                        <span className="detail-value">{teacher.teacher_code}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Email:</span>
                        <span className="detail-value">{getTeacherEmail(teacher)}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Phone:</span>
                        <span className="detail-value">{teacher.phone_number || 'Not provided'}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Specialization:</span>
                        <span className="detail-value">{teacher.specialization || 'Not specified'}</span>
                      </div>
                    </div>

                    <div className="teacher-card-actions">
                      <div className="action-buttons">
                        <Button
                          size="sm"
                          variant="outline"
                          onClick={() => handleViewTeacher(teacher)}
                          fullWidth
                        >
                          View
                        </Button>
                        <Button
                          size="sm"
                          variant="secondary"
                          onClick={() => setEditTeacher({
                            teacher_id: teacher.teacher_id,
                            fullname: getTeacherDisplayName(teacher),
                            email: getTeacherEmail(teacher),
                            teacher_code: teacher.teacher_code,
                            phone_number: teacher.phone_number || '',
                            specialization: teacher.specialization || '',
                            is_active: teacher.is_active
                          })}
                          fullWidth
                        >
                          Edit
                        </Button>
                        {teacher.is_active ? (
                          <Button
                            size="sm"
                            variant="danger"
                            onClick={() => setDeleteTeacher(teacher)}
                            fullWidth
                          >
                            Deactivate
                          </Button>
                        ) : (
                          <Button
                            size="sm"
                            variant="success"
                            onClick={() => setReactivateTeacher(teacher)}
                            fullWidth
                          >
                            Reactivate
                          </Button>
                        )}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </Card>

      {/* Add Teacher Modal */}
      <Modal 
        isOpen={showAddModal} 
        onClose={() => {
          setShowAddModal(false);
          resetForm();
        }} 
        title="Add New Teacher" 
        size="lg"
      >
        <form onSubmit={handleAddSubmit} className="teacher-form">
          {formErrors.submit && <div className="form-error">{formErrors.submit}</div>}

          <div className="form-row">
            <Input 
              label="Full Name *" 
              name="fullname" 
              value={formData.fullname} 
              onChange={handleChange} 
              error={formErrors.fullname} 
              required 
              placeholder="Enter teacher's full name"
            />
            <Input 
              label="Email *" 
              name="email" 
              type="email" 
              value={formData.email} 
              onChange={handleChange} 
              error={formErrors.email} 
              required 
              placeholder="Enter email address"
            />
          </div>

          <div className="form-row">
            <div className="input-with-button">
              <Input 
                label="Teacher Code *" 
                name="teacher_code" 
                value={formData.teacher_code} 
                onChange={handleChange} 
                error={formErrors.teacher_code} 
                required 
                placeholder="Enter unique teacher code"
              />
              <Button type="button" variant="outline" onClick={generateTeacherCode}>
                Generate
              </Button>
            </div>
            <Input 
              label="Phone Number" 
              name="phone_number" 
              value={formData.phone_number} 
              onChange={handleChange} 
              error={formErrors.phone_number} 
              placeholder="Optional phone number"
            />
          </div>

          <div className="form-row">
            <Input 
              label="Specialization" 
              name="specialization" 
              value={formData.specialization} 
              onChange={handleChange} 
              placeholder="e.g., Mathematics, Science"
            />
          </div>

          <div className="form-row">
            <Input 
              label="Password *" 
              name="password" 
              type="password" 
              value={formData.password} 
              onChange={handleChange} 
              error={formErrors.password} 
              required 
              placeholder="Enter password (min. 6 characters)"
            />
            <Input 
              label="Confirm Password *" 
              name="confirmPassword" 
              type="password" 
              value={formData.confirmPassword} 
              onChange={handleChange} 
              error={formErrors.confirmPassword} 
              required 
              placeholder="Confirm your password"
            />
          </div>

          <div className="form-actions">
            <Button type="button" variant="secondary" onClick={() => {
              setShowAddModal(false);
              resetForm();
            }}>
              Cancel
            </Button>
            <Button type="submit" loading={isSubmitting}>
              {isSubmitting ? 'Creating...' : 'Add Teacher'}
            </Button>
          </div>
        </form>
      </Modal>

      {/* Edit Teacher Modal */}
      <Modal isOpen={!!editTeacher} onClose={() => setEditTeacher(null)} title="Edit Teacher" size="lg">
        {editTeacher && (
          <form onSubmit={handleEditSubmit} className="teacher-form">
            {formErrors.submit && <div className="form-error">{formErrors.submit}</div>}

            <div className="form-row">
              <Input 
                label="Full Name *" 
                name="fullname" 
                value={editTeacher.fullname} 
                onChange={handleEditChange} 
                required 
              />
              <Input 
                label="Email *" 
                name="email" 
                type="email" 
                value={editTeacher.email} 
                onChange={handleEditChange} 
                required 
              />
            </div>

            <div className="form-row">
              <Input 
                label="Teacher Code *" 
                name="teacher_code" 
                value={editTeacher.teacher_code} 
                onChange={handleEditChange} 
                required 
              />
              <Input 
                label="Phone Number" 
                name="phone_number" 
                value={editTeacher.phone_number} 
                onChange={handleEditChange} 
              />
            </div>

            <div className="form-row">
              <Input 
                label="Specialization" 
                name="specialization" 
                value={editTeacher.specialization} 
                onChange={handleEditChange} 
              />
            </div>

            <Select 
              label="Status" 
              name="is_active" 
              value={editTeacher.is_active} 
              onChange={(e) => setEditTeacher(prev => ({ ...prev, is_active: e.target.value === 'true' }))}
            >
              <option value={true}>Active</option>
              <option value={false}>Inactive</option>
            </Select>

            <div className="form-actions">
              <Button type="button" variant="secondary" onClick={() => setEditTeacher(null)}>
                Cancel
              </Button>
              <Button type="submit" loading={modalLoading}>
                {modalLoading ? 'Saving...' : 'Save Changes'}
              </Button>
            </div>
          </form>
        )}
      </Modal>

      {/* View Teacher Modal */}
      <Modal isOpen={!!viewTeacher} onClose={() => setViewTeacher(null)} title="Teacher Details" size="xl">
        {viewTeacher && teacherDetails && (
          <div className="teacher-details">
            <div className="detail-section">
              <h3>Basic Information</h3>
              <div className="detail-grid">
                <div className="detail-item">
                  <strong>Full Name:</strong> {getTeacherDisplayName(teacherDetails)}
                </div>
                <div className="detail-item">
                  <strong>Email:</strong> {getTeacherEmail(teacherDetails)}
                </div>
                <div className="detail-item">
                  <strong>Teacher Code:</strong> {teacherDetails.teacher_code}
                </div>
                <div className="detail-item">
                  <strong>Phone:</strong> {teacherDetails.phone_number || 'Not provided'}
                </div>
                <div className="detail-item">
                  <strong>Specialization:</strong> {teacherDetails.specialization || 'Not specified'}
                </div>
                <div className="detail-item">
                  <strong>Status:</strong> 
                  <Badge variant={teacherDetails.is_active ? 'success' : 'danger'}>
                    {teacherDetails.is_active ? 'Active' : 'Inactive'}
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
                    <strong>Total Students:</strong> {stats.student_count || 0}
                  </div>
                </div>
              </div>
            )}

            <div className="detail-section">
              <div className="section-header">
                <h3>Assignments</h3>
                <div className="assignment-filters">
                  <Select 
                    value={assignmentFilters.term_id} 
                    onChange={(e) => handleAssignmentFilterChange('term_id', e.target.value)}
                    style={{ width: '150px' }}
                  >
                    <option value="">All Terms</option>
                    {/* Add terms options here */}
                  </Select>
                  <Input 
                    type="number"
                    placeholder="Academic Year"
                    value={assignmentFilters.academic_year}
                    onChange={(e) => handleAssignmentFilterChange('academic_year', e.target.value)}
                    style={{ width: '120px' }}
                  />
                </div>
              </div>
              <div className="assignments-list">
                {assignments.length === 0 ? (
                  <div className="no-data">No assignments found</div>
                ) : (
                  assignments.map(assignment => (
                    <div key={assignment.assignment_id} className="assignment-item">
                      <strong>{assignment.assignment_subject?.subject_name}</strong>
                      <span> - {assignment.assignment_class?.class_name}</span>
                      {assignment.assignment_stream && (
                        <span> ({assignment.assignment_stream.stream_name})</span>
                      )}
                      <span> - {assignment.assignment_term?.term_name} {assignment.academic_year}</span>
                    </div>
                  ))
                )}
              </div>
            </div>

            <div className="detail-section">
              <h3>Students</h3>
              <div className="students-list">
                {students.length === 0 ? (
                  <div className="no-data">No students assigned</div>
                ) : (
                  students.map(student => (
                    <div key={student.student_id} className="student-item">
                      <strong>{student.fullname}</strong>
                      <span> - {student.student_class?.class_name}</span>
                      {student.student_stream && (
                        <span> ({student.student_stream.stream_name})</span>
                      )}
                      <span> - {student.admission_number}</span>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* Delete Confirmation Modal */}
      <Modal isOpen={!!deleteTeacher} onClose={() => setDeleteTeacher(null)} title="Confirm Deactivation" size="sm">
        {deleteTeacher && (
          <>
            <p>Are you sure you want to deactivate <strong>{getTeacherDisplayName(deleteTeacher)}</strong>?</p>
            <p className="text-muted">This will make the teacher inactive but preserve their data.</p>
            <div className="modal-actions">
              <Button variant="secondary" onClick={() => setDeleteTeacher(null)}>
                Cancel
              </Button>
              <Button variant="danger" onClick={handleDelete} loading={modalLoading}>
                {modalLoading ? 'Deactivating...' : 'Deactivate'}
              </Button>
            </div>
          </>
        )}
      </Modal>

      {/* Reactivate Confirmation Modal */}
      <Modal isOpen={!!reactivateTeacher} onClose={() => setReactivateTeacher(null)} title="Confirm Reactivation" size="sm">
        {reactivateTeacher && (
          <>
            <p>Are you sure you want to reactivate <strong>{getTeacherDisplayName(reactivateTeacher)}</strong>?</p>
            <div className="modal-actions">
              <Button variant="secondary" onClick={() => setReactivateTeacher(null)}>
                Cancel
              </Button>
              <Button variant="success" onClick={handleReactivate} loading={modalLoading}>
                {modalLoading ? 'Reactivating...' : 'Reactivate'}
              </Button>
            </div>
          </>
        )}
      </Modal>
    </div>
  );
};

export default ManageTeachers;