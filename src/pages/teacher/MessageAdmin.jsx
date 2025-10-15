// src/pages/teacher/MessageAdmin.jsx
import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { post } from '../../api/api';
import Card from '../../components/ui/Card';
import Input from '../../components/ui/Input';
import Button from '../../components/ui/Button';
import './MessageAdmin.css';

const MessageAdmin = () => {
  const navigate = useNavigate();
  
  const [formData, setFormData] = useState({
    category: 'Academic',
    subject: '',
    message: ''
  });
  const [formErrors, setFormErrors] = useState({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [success, setSuccess] = useState(false);

  const categories = [
    'Academic',
    'Facilities',
    'Staffing',
    'Student Behavior',
    'Curriculum',
    'Other'
  ];

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
    
    if (formErrors[name]) {
      setFormErrors(prev => ({ ...prev, [name]: '' }));
    }
  };

  const validateForm = () => {
    const errors = {};
    if (!formData.subject.trim()) errors.subject = 'Subject is required';
    if (!formData.message.trim()) errors.message = 'Message is required';
    
    setFormErrors(errors);
    return Object.keys(errors).length === 0;
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!validateForm()) return;

    setIsSubmitting(true);
    setSuccess(false);

    try {
      await post('/comments', {
        category: formData.category,
        subject: formData.subject,
        message: formData.message
      });
      
      setSuccess(true);
      setFormData({ category: 'Academic', subject: '', message: '' });
      
      // Auto-redirect after 3 seconds
      setTimeout(() => {
        navigate('/teacher');
      }, 3000);
    } catch (err) {
      setFormErrors({ submit: err.message || 'Failed to send message. Please try again.' });
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="message-admin">
      <div className="page-header">
        <h1 className="page-title">Message Admin (DOS)</h1>
        <p className="page-subtitle">
          Send a message, suggestion, or complaint to the Director of Studies
        </p>
      </div>

      <Card>
        {success ? (
          <div className="success-message">
            <div className="success-icon">✓</div>
            <h3 className="success-title">Message Sent!</h3>
            <p className="success-text">
              Your message has been sent to the DOS. You will be redirected to your dashboard shortly.
            </p>
          </div>
        ) : (
          <form onSubmit={handleSubmit} className="message-form">
            {formErrors.submit && (
              <div className="form-error">{formErrors.submit}</div>
            )}
            
            <div className="form-field">
              <label htmlFor="category" className="form-label">
                Category <span className="required">*</span>
              </label>
              <select
                id="category"
                name="category"
                value={formData.category}
                onChange={handleChange}
                className="form-select"
                aria-invalid={!!formErrors.category}
              >
                {categories.map(cat => (
                  <option key={cat} value={cat}>{cat}</option>
                ))}
              </select>
              {formErrors.category && (
                <p className="form-error-text">{formErrors.category}</p>
              )}
            </div>
            
            <Input
              label="Subject"
              id="subject"
              name="subject"
              value={formData.subject}
              onChange={handleChange}
              error={formErrors.subject}
              required
            />
            
            <div className="form-field">
              <label htmlFor="message" className="form-label">
                Message <span className="required">*</span>
              </label>
              <textarea
                id="message"
                name="message"
                value={formData.message}
                onChange={handleChange}
                className={`form-textarea ${formErrors.message ? 'form-textarea--error' : ''}`}
                rows="6"
                placeholder="Please provide details about your concern or suggestion..."
                aria-invalid={!!formErrors.message}
              />
              {formErrors.message && (
                <p className="form-error-text">{formErrors.message}</p>
              )}
            </div>
            
            <div className="form-actions">
              <Button
                type="button"
                variant="secondary"
                onClick={() => navigate('/teacher')}
                disabled={isSubmitting}
              >
                Cancel
              </Button>
              <Button
                type="submit"
                variant="primary"
                disabled={isSubmitting}
              >
                {isSubmitting ? 'Sending...' : 'Send Message'}
              </Button>
            </div>
          </form>
        )}
      </Card>
    </div>
  );
};

export default MessageAdmin;