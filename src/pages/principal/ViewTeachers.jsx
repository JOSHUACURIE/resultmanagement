// src/pages/principal/ViewTeachers.jsx
import React, { useState, useEffect, useMemo } from 'react';
import { get } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Table from '../../components/ui/Table';
import Badge from '../../components/ui/Badge';
import './ViewTeachers.css';

const ViewTeachers = () => {
  const [teachers, setTeachers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [subjectFilter, setSubjectFilter] = useState('all');

  // Fetch all teachers
  const fetchTeachers = async () => {
    try {
      const data = await get('/teachers');
      setTeachers(data);
      setError('');
    } catch (err) {
      console.error('Failed to fetch teachers:', err);
      setError('Failed to load teachers. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchTeachers();
  }, []);

  // Get unique subjects for filter dropdown
  const subjects = useMemo(() => {
    const unique = [...new Set(teachers.map(t => t.subject))];
    return unique.sort();
  }, [teachers]);

  // Filter and search teachers
  const filteredTeachers = useMemo(() => {
    return teachers.filter(teacher => {
      // Search
      const matchesSearch = 
        teacher.fullName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        teacher.email.toLowerCase().includes(searchTerm.toLowerCase());
      
      // Subject filter
      const matchesSubject = subjectFilter === 'all' || teacher.subject === subjectFilter;
      
      return matchesSearch && matchesSubject;
    });
  }, [teachers, searchTerm, subjectFilter]);

  // Table columns
  const columns = [
    { key: 'fullName', header: 'Full Name' },
    { 
      key: 'role', 
      header: 'Role',
      render: () => <Badge variant="teacher">Teacher</Badge>
    },
    { key: 'subject', header: 'Subject' },
    { key: 'email', header: 'Email' },
    { 
      key: 'phoneNumber', 
      header: 'Phone',
      render: (phone) => phone || '—'
    },
    { 
      key: 'dateJoined', 
      header: 'Joined',
      render: (date) => date ? new Date(date).toLocaleDateString() : '—'
    }
  ];

  return (
    <div className="view-teachers">
      <div className="page-header">
        <h1 className="page-title">View Teachers</h1>
        <p className="page-subtitle">
          Total teachers: <strong>{teachers.length}</strong>
        </p>
      </div>

      {/* Filters */}
      <Card className="filters-card">
        <div className="filters-grid">
          <Input
            label="Search"
            id="search"
            placeholder="Search by name or email..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="search-input"
          />
          
          <div className="form-field">
            <label htmlFor="subject-filter" className="form-label">
              Subject
            </label>
            <select
              id="subject-filter"
              value={subjectFilter}
              onChange={(e) => setSubjectFilter(e.target.value)}
              className="form-select"
            >
              <option value="all">All Subjects</option>
              {subjects.map(subject => (
                <option key={subject} value={subject}>
                  {subject}
                </option>
              ))}
            </select>
          </div>
        </div>
      </Card>

      {/* Teachers Table */}
      <Card>
        <Table
          columns={columns}
          data={filteredTeachers}
          loading={loading}
          emptyMessage={
            searchTerm || subjectFilter !== 'all'
              ? "No teachers match your search criteria."
              : "No teachers found."
          }
          error={error}
        />
      </Card>
    </div>
  );
};

export default ViewTeachers;