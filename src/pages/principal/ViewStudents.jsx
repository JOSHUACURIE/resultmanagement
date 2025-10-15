// src/pages/principal/ViewStudents.jsx
import React, { useState, useEffect, useMemo } from 'react';
import { get } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Table from '../../components/ui/Table';
import './ViewStudents.css';

const ViewStudents = () => {
  const [students, setStudents] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [classFilter, setClassFilter] = useState('all');

  // Fetch all students
  const fetchStudents = async () => {
    try {
      const data = await get('/students');
      setStudents(data);
      setError('');
    } catch (err) {
      console.error('Failed to fetch students:', err);
      setError('Failed to load students. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchStudents();
  }, []);

  // Get unique classes for filter dropdown
  const classes = useMemo(() => {
    const unique = [...new Set(students.map(s => s.class))];
    return unique.sort();
  }, [students]);

  // Filter and search students
  const filteredStudents = useMemo(() => {
    return students.filter(student => {
      // Search
      const matchesSearch = 
        student.fullName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        student.regNumber.toLowerCase().includes(searchTerm.toLowerCase());
      
      // Class filter
      const matchesClass = classFilter === 'all' || student.class === classFilter;
      
      return matchesSearch && matchesClass;
    });
  }, [students, searchTerm, classFilter]);

  // Table columns
  const columns = [
    { key: 'fullName', header: 'Full Name' },
    { key: 'regNumber', header: 'Reg. No.' },
    { key: 'class', header: 'Class' },
    { 
      key: 'email', 
      header: 'Email',
      render: (email) => email || '—'
    },
    { 
      key: 'phoneNumber', 
      header: 'Phone',
      render: (phone) => phone || '—'
    },
    { 
      key: 'dateOfBirth', 
      header: 'DOB',
      render: (dob) => dob ? new Date(dob).toLocaleDateString() : '—'
    }
  ];

  return (
    <div className="view-students">
      <div className="page-header">
        <h1 className="page-title">View Students</h1>
        <p className="page-subtitle">
          Total students: <strong>{students.length}</strong>
        </p>
      </div>

      {/* Filters */}
      <Card className="filters-card">
        <div className="filters-grid">
          <Input
            label="Search"
            id="search"
            placeholder="Search by name or reg number..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="search-input"
          />
          
          <div className="form-field">
            <label htmlFor="class-filter" className="form-label">
              Class
            </label>
            <select
              id="class-filter"
              value={classFilter}
              onChange={(e) => setClassFilter(e.target.value)}
              className="form-select"
            >
              <option value="all">All Classes</option>
              {classes.map(cls => (
                <option key={cls} value={cls}>
                  {cls}
                </option>
              ))}
            </select>
          </div>
        </div>
      </Card>

      {/* Students Table */}
      <Card>
        <Table
          columns={columns}
          data={filteredStudents}
          loading={loading}
          emptyMessage={
            searchTerm || classFilter !== 'all'
              ? "No students match your search criteria."
              : "No students found."
          }
          error={error}
        />
      </Card>
    </div>
  );
};

export default ViewStudents;