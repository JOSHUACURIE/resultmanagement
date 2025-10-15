// src/pages/dos/ComplaintsSuggestions.jsx
import React, { useState, useEffect } from 'react';
import { get, put } from '../../api/api';
import Card from '../../components/ui/Card';
import Button from '../../components/ui/Button';
import Table from '../../components/ui/Table';
import Modal from '../../components/ui/Modal';
import Badge from '../../components/ui/Badge';
import './ComplaintsSuggestions.css';

const ComplaintsSuggestions = () => {
  const [complaints, setComplaints] = useState([]);
  const [filteredComplaints, setFilteredComplaints] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [selectedComplaint, setSelectedComplaint] = useState(null);
  const [response, setResponse] = useState('');
  const [resolving, setResolving] = useState(false);
  const [filter, setFilter] = useState('all'); // 'all', 'pending', 'resolved'

  // Fetch complaints
  const fetchComplaints = async () => {
    try {
      const data = await get('/complaints');
      setComplaints(data);
      applyFilter(data, filter);
      setError('');
    } catch (err) {
      console.error('Failed to fetch complaints:', err);
      setError('Failed to load complaints. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchComplaints();
  }, []);

  // Apply filter
  const applyFilter = (data, status) => {
    let filtered = data;
    if (status === 'pending') {
      filtered = data.filter(c => c.status === 'Pending');
    } else if (status === 'resolved') {
      filtered = data.filter(c => c.status === 'Resolved');
    }
    setFilteredComplaints(filtered);
  };

  const handleFilterChange = (status) => {
    setFilter(status);
    applyFilter(complaints, status);
  };

  // Mark as resolved
  const handleResolve = async () => {
    if (!selectedComplaint) return;

    setResolving(true);
    try {
      await put(`/complaints/${selectedComplaint.id}/resolve`, { response });
      setSelectedComplaint(null);
      setResponse('');
      fetchComplaints(); // Refresh
    } catch (err) {
      alert('Failed to resolve complaint: ' + err.message);
    } finally {
      setResolving(false);
    }
  };

  // Get category badge variant
  const getCategoryVariant = (category) => {
    const lower = category.toLowerCase();
    if (lower.includes('academic')) return 'info';
    if (lower.includes('behavior') || lower.includes('discipline')) return 'warning';
    if (lower.includes('facility') || lower.includes('infrastructure')) return 'neutral';
    return 'primary';
  };

  // Table columns
  const columns = [
    { 
      key: 'fromName', 
      header: 'From',
      render: (name, row) => (
        <div>
          <div>{name}</div>
          <div className="complaint-role">{row.fromRole}</div>
        </div>
      )
    },
    { 
      key: 'category', 
      header: 'Category',
      render: (category) => (
        <Badge variant={getCategoryVariant(category)}>
          {category}
        </Badge>
      )
    },
    { 
      key: 'message', 
      header: 'Message',
      render: (msg) => (
        <span className="complaint-message">
          {msg.length > 50 ? `${msg.substring(0, 50)}...` : msg}
        </span>
      )
    },
    { 
      key: 'submittedAt', 
      header: 'Date',
      render: (date) => new Date(date).toLocaleDateString()
    },
    { 
      key: 'status', 
      header: 'Status',
      render: (status) => (
        <Badge variant={status === 'Resolved' ? 'success' : 'warning'}>
          {status}
        </Badge>
      )
    },
    {
      key: 'actions',
      header: 'Actions',
      render: (_, complaint) => (
        <Button
          size="sm"
          variant={complaint.status === 'Resolved' ? 'outline' : 'primary'}
          onClick={() => setSelectedComplaint(complaint)}
        >
          {complaint.status === 'Resolved' ? 'View' : 'Respond'}
        </Button>
      )
    }
  ];

  return (
    <div className="complaints-suggestions">
      <div className="page-header">
        <div>
          <h1 className="page-title">Complaints & Suggestions</h1>
          <p className="page-subtitle">Manage feedback from staff, students, and parents</p>
        </div>
        
        <div className="filter-buttons">
          <Button
            variant={filter === 'all' ? 'primary' : 'outline'}
            size="sm"
            onClick={() => handleFilterChange('all')}
          >
            All ({complaints.length})
          </Button>
          <Button
            variant={filter === 'pending' ? 'primary' : 'outline'}
            size="sm"
            onClick={() => handleFilterChange('pending')}
          >
            Pending ({complaints.filter(c => c.status === 'Pending').length})
          </Button>
          <Button
            variant={filter === 'resolved' ? 'primary' : 'outline'}
            size="sm"
            onClick={() => handleFilterChange('resolved')}
          >
            Resolved ({complaints.filter(c => c.status === 'Resolved').length})
          </Button>
        </div>
      </div>

      {/* Complaints Table */}
      <Card>
        <Table
          columns={columns}
          data={filteredComplaints}
          loading={loading}
          emptyMessage="No complaints or suggestions found."
          error={error}
        />
      </Card>

      {/* View/Respond Modal */}
      <Modal
        isOpen={!!selectedComplaint}
        onClose={() => {
          setSelectedComplaint(null);
          setResponse('');
        }}
        title={selectedComplaint?.status === 'Resolved' ? 'Complaint Details' : 'Respond to Complaint'}
        size="md"
      >
        {selectedComplaint && (
          <>
            <div className="complaint-details">
              <div className="complaint-meta">
                <div>
                  <strong>From:</strong> {selectedComplaint.fromName} ({selectedComplaint.fromRole})
                </div>
                <div>
                  <strong>Category:</strong> {selectedComplaint.category}
                </div>
                <div>
                  <strong>Date:</strong> {new Date(selectedComplaint.submittedAt).toLocaleString()}
                </div>
                <div>
                  <strong>Status:</strong> 
                  <Badge 
                    variant={selectedComplaint.status === 'Resolved' ? 'success' : 'warning'}
                    className="ml-2"
                  >
                    {selectedComplaint.status}
                  </Badge>
                </div>
              </div>
              
              <div className="complaint-message-full">
                <strong>Message:</strong>
                <p>{selectedComplaint.message}</p>
              </div>

              {selectedComplaint.response && (
                <div className="complaint-response">
                  <strong>Your Response:</strong>
                  <p>{selectedComplaint.response}</p>
                </div>
              )}
            </div>

            {selectedComplaint.status !== 'Resolved' && (
              <>
                <div className="form-field mt-4">
                  <label htmlFor="response" className="form-label">
                    Your Response (optional)
                  </label>
                  <textarea
                    id="response"
                    value={response}
                    onChange={(e) => setResponse(e.target.value)}
                    className="response-textarea"
                    rows="3"
                    placeholder="Write a response to the complainant..."
                  />
                </div>
                
                <div className="modal-footer">
                  <Button
                    variant="secondary"
                    onClick={() => {
                      setSelectedComplaint(null);
                      setResponse('');
                    }}
                    disabled={resolving}
                  >
                    Cancel
                  </Button>
                  <Button
                    variant="primary"
                    onClick={handleResolve}
                    disabled={resolving}
                  >
                    {resolving ? 'Resolving...' : 'Mark as Resolved'}
                  </Button>
                </div>
              </>
            )}
          </>
        )}
      </Modal>
    </div>
  );
};

export default ComplaintsSuggestions;