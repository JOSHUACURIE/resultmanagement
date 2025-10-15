// src/api/teacherApi.js - FRONTEND ONLY
import { get, post, put, del } from './api';

export const teacherApi = {
  // Assignment methods - FRONTEND API CLIENT
  getMyAssignments: async (academicYear = null, termId = null) => {
    try {
      const params = new URLSearchParams();
      if (academicYear) params.append('academic_year', academicYear);
      if (termId) params.append('term_id', termId);
      
      const queryString = params.toString();
      const endpoint = `/teachers/my-assignments${queryString ? `?${queryString}` : ''}`;
      
      console.log('🔄 Frontend - Calling getMyAssignments:', endpoint);
      return await get(endpoint);
    } catch (error) {
      console.error('❌ Frontend - Error in getMyAssignments:', error);
      throw error;
    }
  },

  getMySubjects: async (academicYear = null, termId = null) => {
    try {
      const params = new URLSearchParams();
      if (academicYear) params.append('academic_year', academicYear);
      if (termId) params.append('term_id', termId);
      
      const queryString = params.toString();
      const endpoint = `/teachers/my-subjects${queryString ? `?${queryString}` : ''}`;
      
      return await get(endpoint);
    } catch (error) {
      console.error('Error in getMySubjects:', error);
      throw error;
    }
  },

  getMyAssignmentDetails: async (assignmentId) => {
    try {
      return await get(`/teachers/my-assignments/${assignmentId}`);
    } catch (error) {
      console.error('Error in getMyAssignmentDetails:', error);
      throw error;
    }
  },

  getMyAssignmentStats: async (academicYear = null, termId = null) => {
    try {
      const params = new URLSearchParams();
      if (academicYear) params.append('academic_year', academicYear);
      if (termId) params.append('term_id', termId);
      
      const queryString = params.toString();
      const endpoint = `/teachers/my-assignment-stats${queryString ? `?${queryString}` : ''}`;
      
      return await get(endpoint);
    } catch (error) {
      console.error('Error in getMyAssignmentStats:', error);
      throw error;
    }
  },

  getMyTimetable: async (academicYear = null, termId = null) => {
    try {
      const params = new URLSearchParams();
      if (academicYear) params.append('academic_year', academicYear);
      if (termId) params.append('term_id', termId);
      
      const queryString = params.toString();
      const endpoint = `/teachers/my-timetable${queryString ? `?${queryString}` : ''}`;
      
      return await get(endpoint);
    } catch (error) {
      console.error('Error in getMyTimetable:', error);
      throw error;
    }
  },

  // Student methods
  getMyStudents: async () => {
    try {
      return await get('/teachers/my-students');
    } catch (error) {
      console.error('Error in getMyStudents:', error);
      throw error;
    }
  },
};