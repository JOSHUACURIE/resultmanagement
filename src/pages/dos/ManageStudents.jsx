import React, { useState, useEffect } from 'react';
import { get, post, del } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Select from '../../components/ui/Select';
import Button from '../../components/ui/Button';
import Modal from '../../components/ui/Modal';
import './ManageStudents.css';

const ManageStudents = () => {
  const [formData, setFormData] = useState({
    admission_number: '',
    fullname: '',
    guardian_phone: '',
    class_id: '',
    stream_id: '',
    date_of_birth: '',
    gender: ''
  });
  const [formErrors, setFormErrors] = useState({});
  const [isSubmitting, setIsSubmitting] = useState(false);

  const [students, setStudents] = useState([]);
  const [classes, setClasses] = useState([]);
  const [streams, setStreams] = useState([]);
  const [teachers, setTeachers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tableError, setTableError] = useState('');

  const [showAddModal, setShowAddModal] = useState(false);
  const [deleteStudent, setDeleteStudent] = useState(null);
  const [modalLoading, setModalLoading] = useState(false);
  const [viewStudent, setViewStudent] = useState(null);
  const [filterClass, setFilterClass] = useState('');
  const [filterStream, setFilterStream] = useState('');

  const genderOptions = [
    { value: 'male', label: 'Male' },
    { value: 'female', label: 'Female' },
    { value: 'other', label: 'Other' }
  ];

  useEffect(() => {
    fetchStudents();
    fetchClasses();
    fetchStreams();
    fetchTeachers();
  }, []);

  useEffect(() => {
    fetchStudents();
  }, [filterClass, filterStream]);

  const fetchStudents = async () => {
    try {
      const params = new URLSearchParams();
      if (filterClass) params.append('class_id', filterClass);
      if (filterStream) params.append('stream_id', filterStream);
      
      const response = await get(`/students?${params.toString()}`);
      const studentsData = response.success ? response.data : 
                          Array.isArray(response) ? response : 
                          response.data ? response.data : [];
      setStudents(studentsData);
      setTableError('');
    } catch (err) {
      console.error('Error fetching students:', err);
      setTableError('Failed to load students.');
      setStudents([]);
    } finally {
      setLoading(false);
    }
  };

  const fetchClasses = async () => {
    try {
      const response = await get('/classes');
      const classesData = response.success ? response.data : 
                         Array.isArray(response) ? response : 
                         response.data ? response.data : [];
      setClasses(classesData);
    } catch (err) {
      console.error('Error fetching classes:', err);
      setClasses([]);
    }
  };

  const fetchStreams = async () => {
    try {
      const response = await get('/streams');
      const streamsData = response.success ? response.data : 
                         Array.isArray(response) ? response : 
                         response.data ? response.data : [];
      setStreams(streamsData);
    } catch (err) {
      console.error('Error fetching streams:', err);
      setStreams([]);
    }
  };

  const fetchTeachers = async () => {
    try {
      const response = await get('/teachers');
      const teachersData = response.success ? response.data : 
                          Array.isArray(response) ? response : 
                          response.data ? response.data : [];
      setTeachers(teachersData);
    } catch (err) {
      console.error('Error fetching teachers:', err);
      setTeachers([]);
    }
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
    if (formErrors[name]) setFormErrors(prev => ({ ...prev, [name]: '' }));
  };

  const handleFilterChange = (e) => {
    const { name, value } = e.target;
    if (name === 'filterClass') {
      setFilterClass(value);
    } else if (name === 'filterStream') {
      setFilterStream(value);
    }
  };

  const validateForm = () => {
    const errors = {};
    if (!formData.admission_number.trim()) errors.admission_number = 'Admission number is required';
    if (!formData.fullname.trim()) errors.fullname = 'Full name is required';
    if (!formData.class_id) errors.class_id = 'Class is required';
    
    setFormErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleAddSubmit = async (e) => {
    e.preventDefault();
    if (!validateForm()) return;

    setIsSubmitting(true);
    try {
      const studentData = {
        admission_number: formData.admission_number,
        fullname: formData.fullname,
        guardian_phone: formData.guardian_phone || null,
        class_id: formData.class_id ? parseInt(formData.class_id) : null,
        stream_id: formData.stream_id ? parseInt(formData.stream_id) : null,
        date_of_birth: formData.date_of_birth || null,
        gender: formData.gender || null, 
      };
      
      console.log('Sending student data:', studentData); 
      
      const response = await post('/students/', studentData);
      
      if (response.success) {
        resetForm();
        setShowAddModal(false);
        fetchStudents();
        alert('Student added successfully!');
      } else {
        setFormErrors({ submit: response.message || 'Failed to add student' });
      }
    } catch (err) {
      console.error('Error adding student:', err);
      const errorMessage = err.response?.data?.message || err.response?.data?.error || 'Failed to add student';
      setFormErrors({ submit: errorMessage });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleDelete = async () => {
    if (!deleteStudent) return;
    setModalLoading(true);
    try {
      const response = await del(`/students/${deleteStudent.student_id}`);
      if (response.success) {
        setDeleteStudent(null);
        fetchStudents();
        alert('Student deactivated successfully!');
      } else {
        alert(response.message || 'Failed to deactivate student');
      }
    } catch (err) {
      alert(err.response?.data?.message || 'Failed to deactivate student');
    } finally {
      setModalLoading(false);
    }
  };

  const handleReactivate = async (student) => {
    if (!window.confirm(`Are you sure you want to activate ${student.fullname}?`)) return;
    
    try {
      const response = await post(`/students/${student.student_id}/reactivate`);
      if (response.success) {
        fetchStudents();
        alert('Student activated successfully!');
      } else {
        alert(response.message || 'Failed to activate student');
      }
    } catch (err) {
      alert(err.response?.data?.message || 'Failed to activate student');
    }
  };

  const resetForm = () => {
    setFormData({
      admission_number: '',
      fullname: '',
      guardian_phone: '',
      class_id: '',
      stream_id: '',
      date_of_birth: '',
      gender: ''
    });
    setFormErrors({});
  };

  const getClassStreamName = (student) => {
    const classInfo = student.student_class;
    const streamInfo = student.student_stream;
    
    if (classInfo && streamInfo) {
      return `${classInfo.class_name} - ${streamInfo.stream_name}`;
    } else if (classInfo) {
      return classInfo.class_name;
    }
    return '-';
  };

  const getClassTeacherName = (student) => {
    const teacherInfo = student.class_teacher;
    if (teacherInfo && teacherInfo.user_account) {
      return teacherInfo.user_account.fullname;
    }
    return '-';
  };

  const getStudentScoresSummary = (student) => {
    if (student.student_scores && student.student_scores.length > 0) {
      return `${student.student_scores.length} score(s)`;
    }
    return 'No scores';
  };

  const getGenderLabel = (genderValue) => {
    if (!genderValue) return '-';
    
    const option = genderOptions.find(opt => opt.value === genderValue);
    return option ? option.label : genderValue;
  };

  // Loading state
  if (loading) {
    return (
      <div className="manage-students">
        <div className="page-header">
          <h1 className="page-title">Manage Students</h1>
        </div>
        <Card>
          <div className="loading-state">Loading students...</div>
        </Card>
      </div>
    );
  }

  const safeClasses = Array.isArray(classes) ? classes : [];
  const safeStreams = Array.isArray(streams) ? streams : [];
  const safeTeachers = Array.isArray(teachers) ? teachers : [];

  return (
    <div className="manage-students">
      <div className="page-header">
        <h1 className="page-title">Manage Students</h1>
        <Button onClick={() => setShowAddModal(true)} className="add-student-btn">
          + Add Student
        </Button>
      </div>

      <Card className="filters-card">
        <div className="filters">
          <Select 
            label="Filter by Class" 
            name="filterClass" 
            value={filterClass} 
            onChange={handleFilterChange}
          >
            <option value="">All Classes</option>
            {safeClasses.map(c => (
              <option key={c.class_id} value={c.class_id}>
                {c.class_name}
              </option>
            ))}
          </Select>

          <Select 
            label="Filter by Stream" 
            name="filterStream" 
            value={filterStream} 
            onChange={handleFilterChange}
          >
            <option value="">All Streams</option>
            {safeStreams.map(s => (
              <option key={s.stream_id} value={s.stream_id}>
                {s.stream_name}
              </option>
            ))}
          </Select>

          <Button 
            variant="secondary" 
            onClick={() => {
              setFilterClass('');
              setFilterStream('');
            }}
            className="clear-filters-btn"
          >
            Clear Filters
          </Button>
        </div>
      </Card>

      <Card>
        {tableError && (
          <div className="error-message">
            {tableError}
            <Button variant="secondary" size="sm" onClick={fetchStudents} style={{ marginLeft: '10px' }}>
              Retry
            </Button>
          </div>
        )}
        
        {/* Students Table */}
        <div className="students-table-container">
          {students.length === 0 ? (
            <div className="empty-state">
              <p>No students found.</p>
            </div>
          ) : (
            <>
              {/* Desktop Table */}
              <table className="students-table desktop-table">
                <thead>
                  <tr>
                    <th>Admission No.</th>
                    <th>Full Name</th>
                    <th>Class & Stream</th>
                    <th>Guardian Phone</th>
                    <th>Gender</th>
                    <th>Status</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {students.map((student) => (
                    <tr key={student.student_id} className="student-row">
                      <td>
                        <div className="student-admission">{student.admission_number}</div>
                      </td>
                      <td>
                        <div className="student-name">{student.fullname}</div>
                      </td>
                      <td>
                        <div className="class-stream">{getClassStreamName(student)}</div>
                      </td>
                      <td>
                        <div className="guardian-phone">{student.guardian_phone || '-'}</div>
                      </td>
                      <td>
                        <div className="student-gender">{getGenderLabel(student.gender)}</div>
                      </td>
                      <td>
                        <span className={`status-badge ${student.is_active ? 'active' : 'inactive'}`}>
                          {student.is_active ? 'Active' : 'Inactive'}
                        </span>
                      </td>
                      <td>
                        <div className="action-buttons">
                          <Button 
                            size="sm" 
                            variant="outline" 
                            onClick={() => setViewStudent(student)}
                          >
                            View
                          </Button>
                          {student.is_active ? (
                            <Button 
                              size="sm" 
                              variant="danger" 
                              onClick={() => setDeleteStudent(student)}
                            >
                              Deactivate
                            </Button>
                          ) : (
                            <Button 
                              size="sm" 
                              variant="success" 
                              onClick={() => handleReactivate(student)}
                            >
                              Activate
                            </Button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Mobile Cards */}
              <div className="mobile-students-list">
                {students.map((student) => (
                  <div key={student.student_id} className="student-card">
                    <div className="student-card-header">
                      <div className="student-card-main">
                        <div className="student-card-name">{student.fullname}</div>
                        <div className="student-card-admission">{student.admission_number}</div>
                      </div>
                      <div className="student-card-status">
                        <span className={`status-badge ${student.is_active ? 'active' : 'inactive'}`}>
                          {student.is_active ? 'Active' : 'Inactive'}
                        </span>
                      </div>
                    </div>
                    
                    <div className="student-card-details">
                      <div className="detail-item">
                        <span className="detail-label">Class & Stream:</span>
                        <span className="detail-value">{getClassStreamName(student)}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Guardian Phone:</span>
                        <span className="detail-value">{student.guardian_phone || '-'}</span>
                      </div>
                      <div className="detail-item">
                        <span className="detail-label">Gender:</span>
                        <span className="detail-value">{getGenderLabel(student.gender)}</span>
                      </div>
                    </div>

                    <div className="student-card-actions">
                      <div className="action-buttons">
                        <Button
                          size="sm"
                          variant="outline"
                          onClick={() => setViewStudent(student)}
                          fullWidth
                        >
                          View Details
                        </Button>
                        {student.is_active ? (
                          <Button
                            size="sm"
                            variant="danger"
                            onClick={() => setDeleteStudent(student)}
                            fullWidth
                          >
                            Deactivate
                          </Button>
                        ) : (
                          <Button
                            size="sm"
                            variant="success"
                            onClick={() => handleReactivate(student)}
                            fullWidth
                          >
                            Activate
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

      {/* Add Student Modal */}
      <Modal 
        isOpen={showAddModal} 
        onClose={() => {
          setShowAddModal(false);
          resetForm();
        }} 
        title="Add New Student" 
        size="md"
      >
        <form onSubmit={handleAddSubmit} className="student-form">
          {formErrors.submit && <div className="form-error">{formErrors.submit}</div>}

          <div className="form-row">
            <Input 
              label="Admission Number *" 
              name="admission_number" 
              value={formData.admission_number} 
              onChange={handleChange} 
              error={formErrors.admission_number} 
              required 
            />
            <Input 
              label="Full Name *" 
              name="fullname" 
              value={formData.fullname} 
              onChange={handleChange} 
              error={formErrors.fullname} 
              placeholder="Enter full name"
              required 
            />
          </div>

          <div className="form-row">
            <Select 
              label="Class *" 
              name="class_id" 
              value={formData.class_id} 
              onChange={handleChange}
              error={formErrors.class_id}
              required
            >
              <option value="">-- Select Class --</option>
              {safeClasses.map(c => (
                <option key={c.class_id} value={c.class_id}>
                  {c.class_name}
                </option>
              ))}
            </Select>

            <Select 
              label="Stream" 
              name="stream_id" 
              value={formData.stream_id} 
              onChange={handleChange}
            >
              <option value="">-- Select Stream --</option>
              {safeStreams.map(stream => (
                <option key={stream.stream_id} value={stream.stream_id}>
                  {stream.stream_name}
                </option>
              ))}
            </Select>
          </div>

          <div className="form-row">
            <Input 
              label="Guardian Phone" 
              name="guardian_phone" 
              value={formData.guardian_phone} 
              onChange={handleChange} 
              placeholder="Phone number"
            />
            <Select 
              label="Gender" 
              name="gender" 
              value={formData.gender} 
              onChange={handleChange}
            >
              <option value="">-- Select Gender --</option>
              {genderOptions.map(option => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </Select>
          </div>

          <div className="form-row">
            <Input 
              label="Date of Birth" 
              name="date_of_birth" 
              type="date"
              value={formData.date_of_birth} 
              onChange={handleChange} 
            />
          </div>

          <div className="form-actions">
            <Button 
              type="button" 
              variant="secondary" 
              onClick={() => {
                setShowAddModal(false);
                resetForm();
              }}
            >
              Cancel
            </Button>
            <Button type="submit" loading={isSubmitting}>Add Student</Button>
          </div>
        </form>
      </Modal>

      {/* View Student Modal */}
      <Modal 
        isOpen={!!viewStudent} 
        onClose={() => setViewStudent(null)} 
        title="Student Details" 
        size="md"
      >
        {viewStudent && (
          <div className="student-details">
            <div className="detail-section">
              <div className="detail-grid">
                <div className="detail-item">
                  <strong>Admission No:</strong> {viewStudent.admission_number}
                </div>
                <div className="detail-item">
                  <strong>Full Name:</strong> {viewStudent.fullname}
                </div>
                <div className="detail-item">
                  <strong>Class & Stream:</strong> {getClassStreamName(viewStudent)}
                </div>
                <div className="detail-item">
                  <strong>Class Teacher:</strong> {getClassTeacherName(viewStudent)}
                </div>
                {viewStudent.guardian_phone && (
                  <div className="detail-item">
                    <strong>Guardian Phone:</strong> {viewStudent.guardian_phone}
                  </div>
                )}
                {viewStudent.gender && (
                  <div className="detail-item">
                    <strong>Gender:</strong> {getGenderLabel(viewStudent.gender)}
                  </div>
                )}
                {viewStudent.date_of_birth && (
                  <div className="detail-item">
                    <strong>Date of Birth:</strong> {new Date(viewStudent.date_of_birth).toLocaleDateString()}
                  </div>
                )}
                <div className="detail-item">
                  <strong>Status:</strong> 
                  <span className={`status-badge ${viewStudent.is_active ? 'active' : 'inactive'}`}>
                    {viewStudent.is_active ? 'Active' : 'Inactive'}
                  </span>
                </div>
              </div>
            </div>
          </div>
        )}
      </Modal>

      {/* Delete Confirmation Modal */}
      <Modal 
        isOpen={!!deleteStudent} 
        onClose={() => setDeleteStudent(null)} 
        title="Confirm Deactivation"
      >
        <p>
          Are you sure you want to deactivate {deleteStudent?.fullname}? 
          This will make the student inactive but preserve their data.
        </p>
        <div className="modal-actions">
          <Button variant="secondary" onClick={() => setDeleteStudent(null)}>Cancel</Button>
          <Button variant="danger" onClick={handleDelete} loading={modalLoading}>
            Deactivate
          </Button>
        </div>
      </Modal>
    </div>
  );
};

export default ManageStudents;