import React, { useState, useEffect } from 'react';
import { get, post, del } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Button from '../../components/ui/Button';
import Modal from '../../components/ui/Modal';
import './ManageClasses.css';

const ManageClasses = () => {
  // ==========================
  // STATE
  // ==========================
  const [formData, setFormData] = useState({
    class_name: '',
    class_level: '',
    stream_name: '',
    class_id: '',
  });
  const [formErrors, setFormErrors] = useState({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [classes, setClasses] = useState([]);
  const [streams, setStreams] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tableError, setTableError] = useState('');

  const [showAddClassModal, setShowAddClassModal] = useState(false);
  const [showAddStreamModal, setShowAddStreamModal] = useState(false);
  const [deleteItem, setDeleteItem] = useState(null);
  const [modalLoading, setModalLoading] = useState(false);

  // ==========================
  // FETCH FUNCTIONS
  // ==========================
  const fetchClasses = async () => {
    try {
      const res = await get('/classes');
      const data = Array.isArray(res?.data) ? res.data : res?.data?.data || [];
      setClasses(data);
      setTableError('');
    } catch (err) {
      console.error('❌ Failed to fetch classes:', err);
      setTableError('Failed to load classes. Please try again.');
      setClasses([]);
    } finally {
      setLoading(false);
    }
  };

  const fetchStreams = async () => {
    try {
      const res = await get('/streams');
      const data = Array.isArray(res?.data) ? res.data : res?.data?.data || [];
      setStreams(data);
    } catch (err) {
      console.error('❌ Failed to fetch streams:', err);
      setStreams([]);
    }
  };

  useEffect(() => {
    fetchClasses();
    fetchStreams();
  }, []);

  // ==========================
  // FORM HANDLERS
  // ==========================
  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
    if (formErrors[name]) setFormErrors((prev) => ({ ...prev, [name]: '' }));
  };

  const validateForm = (isStream = false) => {
    const errors = {};
    if (!isStream && !formData.class_name.trim())
      errors.class_name = 'Class name is required';
    if (!isStream && !formData.class_level.trim())
      errors.class_level = 'Class level is required';
    if (isStream && !formData.class_id)
      errors.class_id = 'Select a class first';
    if (isStream && !formData.stream_name.trim())
      errors.stream_name = 'Stream name is required';
    setFormErrors(errors);
    return Object.keys(errors).length === 0;
  };

  // ==========================
  // SUBMIT HANDLERS
  // ==========================
  const handleAddClassSubmit = async (e) => {
    e.preventDefault();
    if (!validateForm()) return;
    setIsSubmitting(true);

    try {
      const res = await post('/classes', {
        class_name: formData.class_name,
        class_level: formData.class_level,
      });

      if (!res?.data?.success && !res?.success) {
        throw new Error(res?.data?.message || 'Failed to add class');
      }

      alert('✅ Class added successfully');
      setFormData({ class_name: '', class_level: '', stream_name: '', class_id: '' });
      setShowAddClassModal(false);
      fetchClasses();
    } catch (err) {
      console.error('❌ Add class error:', err);
      setFormErrors({ submit: err.message || 'Failed to add class' });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleAddStreamSubmit = async (e) => {
    e.preventDefault();
    if (!validateForm(true)) return;
    setIsSubmitting(true);

    try {
      const res = await post('/streams', {
        stream_name: formData.stream_name,
        class_id: formData.class_id,
      });

      if (!res?.data?.success && !res?.success) {
        throw new Error(res?.data?.message || 'Failed to add stream');
      }

      alert('✅ Stream added successfully');
      setFormData({ class_name: '', class_level: '', stream_name: '', class_id: '' });
      setShowAddStreamModal(false);
      fetchStreams();
    } catch (err) {
      console.error('❌ Add stream error:', err);
      setFormErrors({ submit: err.message || 'Failed to add stream' });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleDelete = async () => {
    if (!deleteItem) return;
    setModalLoading(true);

    try {
      if (deleteItem.type === 'class') await del(`/classes/${deleteItem.id}`);
      if (deleteItem.type === 'stream') await del(`/streams/${deleteItem.id}`);

      alert('🗑️ Deleted successfully');
      setDeleteItem(null);
      fetchClasses();
      fetchStreams();
    } catch (err) {
      console.error('❌ Delete error:', err);
      alert('Failed to delete: ' + err.message);
    } finally {
      setModalLoading(false);
    }
  };

  // Loading state
  if (loading) {
    return (
      <div className="manage-classes">
        <div className="page-header">
          <h1 className="page-title">Manage Classes & Streams</h1>
        </div>
        <Card>
          <div className="loading-state">Loading classes and streams...</div>
        </Card>
      </div>
    );
  }

  const safeClasses = Array.isArray(classes) ? classes : [];
  const safeStreams = Array.isArray(streams) ? streams : [];

  return (
    <div className="manage-classes">
      <div className="page-header">
        <h1 className="page-title">Manage Classes & Streams</h1>
        <div className="header-actions">
          <Button 
            onClick={() => setShowAddClassModal(true)} 
            className="add-class-btn"
          >
            + Add Class
          </Button>
          <Button
            onClick={() => {
              if (!Array.isArray(classes) || classes.length === 0) {
                alert('Please add a class first before adding streams.');
                return;
              }
              setShowAddStreamModal(true);
            }}
            className="add-stream-btn"
          >
            + Add Stream
          </Button>
        </div>
      </div>

      {/* CLASSES CARD */}
      <Card>
        <div className="section-header">
          <h3>Classes</h3>
          <span className="item-count">{safeClasses.length} classes</span>
        </div>
        
        <div className="classes-table-container">
          {safeClasses.length === 0 ? (
            <div className="empty-state">
              <p>No classes found. Add your first class!</p>
            </div>
          ) : (
            <>
              {/* Desktop Classes Table */}
              <table className="classes-table desktop-table">
                <thead>
                  <tr>
                    <th>Class Name</th>
                    <th>Level</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {safeClasses.map((cls) => (
                    <tr key={cls.class_id} className="class-row">
                      <td>
                        <div className="class-name">{cls.class_name}</div>
                      </td>
                      <td>
                        <div className="class-level">{cls.class_level}</div>
                      </td>
                      <td>
                        <div className="action-buttons">
                          <Button
                            size="sm"
                            variant="danger"
                            onClick={() =>
                              setDeleteItem({ id: cls.class_id, type: 'class', name: cls.class_name })
                            }
                          >
                            Delete
                          </Button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Mobile Classes List */}
              <div className="mobile-classes-list">
                {safeClasses.map((cls) => (
                  <div key={cls.class_id} className="class-card">
                    <div className="class-card-header">
                      <div className="class-card-main">
                        <div className="class-card-name">{cls.class_name}</div>
                        <div className="class-card-level">Level {cls.class_level}</div>
                      </div>
                    </div>
                    <div className="class-card-actions">
                      <Button
                        size="sm"
                        variant="danger"
                        onClick={() =>
                          setDeleteItem({ id: cls.class_id, type: 'class', name: cls.class_name })
                        }
                        fullWidth
                      >
                        Delete Class
                      </Button>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </Card>

      {/* STREAMS CARD */}
      <Card>
        <div className="section-header">
          <h3>Streams</h3>
          <span className="item-count">{safeStreams.length} streams</span>
        </div>
        
        <div className="streams-table-container">
          {safeStreams.length === 0 ? (
            <div className="empty-state">
              <p>No streams found. Add your first stream!</p>
            </div>
          ) : (
            <>
              {/* Desktop Streams Table */}
              <table className="streams-table desktop-table">
                <thead>
                  <tr>
                    <th>Stream Name</th>
                    <th>Class</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {safeStreams.map((stream) => (
                    <tr key={stream.stream_id} className="stream-row">
                      <td>
                        <div className="stream-name">{stream.stream_name}</div>
                      </td>
                      <td>
                        <div className="stream-class">
                          {safeClasses.find((c) => c.class_id === stream.class_id)?.class_name || '-'}
                        </div>
                      </td>
                      <td>
                        <div className="action-buttons">
                          <Button
                            size="sm"
                            variant="danger"
                            onClick={() =>
                              setDeleteItem({
                                id: stream.stream_id,
                                type: 'stream',
                                name: stream.stream_name,
                              })
                            }
                          >
                            Delete
                          </Button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Mobile Streams List */}
              <div className="mobile-streams-list">
                {safeStreams.map((stream) => (
                  <div key={stream.stream_id} className="stream-card">
                    <div className="stream-card-header">
                      <div className="stream-card-main">
                        <div className="stream-card-name">{stream.stream_name}</div>
                        <div className="stream-card-class">
                          {safeClasses.find((c) => c.class_id === stream.class_id)?.class_name || '-'}
                        </div>
                      </div>
                    </div>
                    <div className="stream-card-actions">
                      <Button
                        size="sm"
                        variant="danger"
                        onClick={() =>
                          setDeleteItem({
                            id: stream.stream_id,
                            type: 'stream',
                            name: stream.stream_name,
                          })
                        }
                        fullWidth
                      >
                        Delete Stream
                      </Button>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </Card>

      {/* ADD CLASS MODAL */}
      <Modal
        isOpen={showAddClassModal}
        onClose={() => setShowAddClassModal(false)}
        title="Add New Class"
        size="md"
      >
        <form onSubmit={handleAddClassSubmit} className="class-form">
          {formErrors.submit && <div className="form-error">{formErrors.submit}</div>}

          <Input
            label="Class Name"
            id="class_name"
            name="class_name"
            value={formData.class_name}
            onChange={handleChange}
            error={formErrors.class_name}
            required
            placeholder="e.g., SS1, JS2"
          />

          <Input
            label="Class Level"
            id="class_level"
            name="class_level"
            value={formData.class_level}
            onChange={handleChange}
            error={formErrors.class_level}
            required
            placeholder="e.g., 1, 2, 3, 4"
          />

          <div className="modal-actions">
            <Button
              type="button"
              variant="secondary"
              onClick={() => setShowAddClassModal(false)}
            >
              Cancel
            </Button>
            <Button type="submit" variant="primary" disabled={isSubmitting}>
              {isSubmitting ? 'Adding...' : 'Add Class'}
            </Button>
          </div>
        </form>
      </Modal>

      {/* ADD STREAM MODAL */}
      <Modal
        isOpen={showAddStreamModal}
        onClose={() => setShowAddStreamModal(false)}
        title="Add New Stream"
        size="md"
      >
        <form onSubmit={handleAddStreamSubmit} className="class-form">
          {formErrors.submit && <div className="form-error">{formErrors.submit}</div>}

          <Input
            label="Stream Name"
            id="stream_name"
            name="stream_name"
            value={formData.stream_name}
            onChange={handleChange}
            error={formErrors.stream_name}
            required
            placeholder="e.g., Science, Arts"
          />

          <div className="form-field">
            <label htmlFor="class_id" className="form-label">Select Class *</label>
            <select
              id="class_id"
              name="class_id"
              value={formData.class_id}
              onChange={handleChange}
              className="form-select"
              required
            >
              <option value="">-- Select Class --</option>
              {Array.isArray(classes) &&
                classes.map((cls) => (
                  <option key={cls.class_id} value={cls.class_id}>
                    {cls.class_name}
                  </option>
                ))}
            </select>
            {formErrors.class_id && (
              <div className="form-error">{formErrors.class_id}</div>
            )}
          </div>

          <div className="modal-actions">
            <Button
              type="button"
              variant="secondary"
              onClick={() => setShowAddStreamModal(false)}
            >
              Cancel
            </Button>
            <Button type="submit" variant="primary" disabled={isSubmitting}>
              {isSubmitting ? 'Adding...' : 'Add Stream'}
            </Button>
          </div>
        </form>
      </Modal>

      {/* DELETE CONFIRMATION MODAL */}
      <Modal
        isOpen={!!deleteItem}
        onClose={() => setDeleteItem(null)}
        title="Confirm Deletion"
        size="sm"
      >
        {deleteItem && (
          <>
            <p>
              Are you sure you want to delete <strong>{deleteItem.name}</strong>? This
              action cannot be undone.
            </p>
            <div className="modal-actions">
              <Button
                variant="secondary"
                onClick={() => setDeleteItem(null)}
                disabled={modalLoading}
              >
                Cancel
              </Button>
              <Button
                variant="danger"
                onClick={handleDelete}
                disabled={modalLoading}
              >
                {modalLoading ? 'Deleting...' : 'Delete'}
              </Button>
            </div>
          </>
        )}
      </Modal>
    </div>
  );
};

export default ManageClasses;