// components/TeacherAssignment.jsx
import React from 'react';
import Badge from './ui/Badge';

const TeacherAssignment = ({ teacher, subjectId, onRemoveAssignment }) => {
  return (
    <div className="teacher-assignment">
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
              onClick={() => onRemoveAssignment(subjectId, cls.assignment_id)}
              title="Remove assignment"
            >
              ×
            </button>
          </Badge>
        ))}
      </div>
    </div>
  );
};

export default TeacherAssignment;