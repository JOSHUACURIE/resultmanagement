// src/pages/principal/ViewComplaints.jsx
import React, { useState, useEffect, useMemo } from 'react';
import { get } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Table from '../../components/ui/Table';
import Modal from '../../components/ui/Modal';
import Badge from '../../components/ui/Badge';
import Button from '../../components/ui/Button';
import './ViewComplaints.css';

const ViewComplaints = () => {
  const [complaints, setComplaints] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('all'); // 'all', 'pending', 'resolved'
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [selectedComplaint, setSelectedComplaint] = useState(null);

  // Fetch all complaints
  const fetchComplaints = async () => {
    try {
      const data = await get('/complaints/all');
      setComplaints(data);
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

  // Get unique categories
  const categories = useMemo(() => {
    const unique = [...new Set(complaints.map(c => c.category))];
    return unique.sort();
  }, [complaints]);

  // Filter complaints
  const filteredComplaints = useMemo(() => {
    return complaints.filter(complaint => {
      // Search
      const matchesSearch = 
        complaint.fromName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        complaint.message.toLowerCase().includes(searchTerm.toLowerCase());
      
      // Status filter
      const matchesStatus = statusFilter === 'all' || complaint.status.toLowerCase() === statusFilter;
      
      // Category filter
      const matchesCategory = categoryFilter === 'all' || complaint.category === categoryFilter;
      
      return matchesSearch && matchesStatus && matchesCategory;
    });
  }, [complaints, searchTerm, statusFilter, categoryFilter]);

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
          variant="outline"
          onClick={() => setSelectedComplaint(complaint)}
        >
          View
        </Button>
      )
    }
  ];

  return (
    <div className="view-complaints">
      <div className="page-header">
        <h1 className="page-title">View Complaints</h1>
        <p className="page-subtitle">
          Total complaints: <strong>{complaints.length}</strong> • 
          Pending: <strong>{complaints.filter(c => c.status === 'Pending').length}</strong>
        </p>
      </div>

      {/* Filters */}
      <Card className="filters-card">
        <div className="filters-grid">
          <Input
            label="Search"
            id="search"
            placeholder="Search by name or message..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="search-input"
          />
          
          <div className="form-field">
            <label htmlFor="status-filter" className="form-label">
              Status
            </label>
            <select
              id="status-filter"
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              className="form-select"
            >
              <option value="all">All Statuses</option>
              <option value="pending">Pending</option>
              <option value="resolved">Resolved</option>
            </select>
          </div>
          
          <div className="form-field">
            <label htmlFor="category-filter" className="form-label">
              Category
            </label>
            <select
              id="category-filter"
              value={categoryFilter}
              onChange={(e) => setCategoryFilter(e.target.value)}
              className="form-select"
            >
              <option value="all">All Categories</option>
              {categories.map(cat => (
                <option key={cat} value={cat}>
                  {cat}
                </option>
              ))}
            </select>
          </div>
        </div>
      </Card>

      {/* Complaints Table */}
      <Card>
        <Table
          columns={columns}
          data={filteredComplaints}
          loading={loading}
          emptyMessage={
            searchTerm || statusFilter !== 'all' || categoryFilter !== 'all'
              ? "No complaints match your filters."
              : "No complaints found."
          }
          error={error}
        />
      </Card>

      {/* View Complaint Modal */}
      <Modal
        isOpen={!!selectedComplaint}
        onClose={() => setSelectedComplaint(null)}
        title="Complaint Details"
        size="md"
      >
        {selectedComplaint && (
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
                <strong>Response:</strong>
                <p>{selectedComplaint.response}</p>
              </div>
            )}
          </div>
        )}
      </Modal>
    </div>
  );
};

export default ViewComplaints;