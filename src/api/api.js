// src/api/api.js - Complete Updated Version

// === Configuration ===
const getAPIBaseURL = () => {
  // Priority 1: Environment variable (Vite)
  if (import.meta.env.VITE_API_BASE_URL) {
    return import.meta.env.VITE_API_BASE_URL;
  }
  
  // Priority 2: Render environment variable
  if (import.meta.env.VITE_RENDER_API_URL) {
    return import.meta.env.VITE_RENDER_API_URL;
  }
  
  // Priority 3: Auto-detect based on current host
  const currentHost = window.location.hostname;
  const isLocalhost = currentHost === 'localhost' || currentHost === '127.0.0.1';
  const isVercel = currentHost.includes('vercel.app');
  
  if (isLocalhost) {
    return 'http://localhost:5000/api';
  } else if (isVercel) {
    return 'https://result-7.onrender.com/api';
  }
  
  // Fallback
  return 'https://result-7.onrender.com/api';
};

const API_BASE_URL = getAPIBaseURL();

console.log('🚀 API Configuration:', {
  baseURL: API_BASE_URL,
  currentHost: window.location.host,
  environment: import.meta.env.MODE,
  envVars: {
    VITE_API_BASE_URL: import.meta.env.VITE_API_BASE_URL,
    VITE_RENDER_API_URL: import.meta.env.VITE_RENDER_API_URL
  }
});

// === Helper Functions ===
const getStoredUser = () => {
  try {
    const raw = localStorage.getItem('user');
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    console.error('❌ Failed to parse user from localStorage:', error);
    return null;
  }
};

const getAuthToken = () => {
  return localStorage.getItem('token');
};

const clearAuthData = () => {
  localStorage.removeItem('user');
  localStorage.removeItem('token');
  console.log('🔐 Auth data cleared');
};

// === Enhanced API Client with Timeout ===
const api = async (endpoint, options = {}) => {
  const url = `${API_BASE_URL}${endpoint}`;
  
  // Create abort controller for timeout
  const controller = new AbortController();
  const timeoutId = setTimeout(() => {
    controller.abort();
    console.log('⏰ API request timed out');
  }, 30000); // 30 second timeout

  // Enhanced logging
  console.log('🔄 API Request:', {
    url,
    method: options.method || 'GET',
    endpoint,
    hasBody: !!options.body,
    timestamp: new Date().toISOString()
  });

  // Prepare headers
  const headers = {
    'Content-Type': 'application/json',
    ...options.headers,
  };

  // Add authorization token
  const token = getAuthToken();
  if (token) {
    headers['Authorization'] = `Bearer ${token}`;
    console.log('🔐 Using auth token: Bearer ***' + token.slice(-8));
  } else {
    console.warn('⚠️ No auth token found for request');
  }

  // Prepare config
  const config = {
    method: options.method || 'GET',
    headers,
    signal: controller.signal,
    credentials: 'include',
  };

  // Handle request body
  if (options.body && !['GET', 'HEAD'].includes(config.method)) {
    if (options.body instanceof FormData) {
      delete headers['Content-Type'];
      config.body = options.body;
    } else {
      config.body = JSON.stringify(options.body);
    }
  }

  try {
    const response = await fetch(url, config);
    clearTimeout(timeoutId);

    // Enhanced response logging
    console.log('📡 API Response:', {
      status: response.status,
      statusText: response.statusText,
      ok: response.ok,
      url: response.url,
      contentType: response.headers.get('content-type'),
      method: config.method
    });

    // Handle response content type
    const contentType = response.headers.get('content-type');
    let data;

    if (contentType && contentType.includes('application/json')) {
      try {
        data = await response.json();
      } catch (parseError) {
        console.error('❌ Failed to parse JSON response:', parseError);
        data = { error: 'Invalid JSON response' };
      }
    } else {
      data = await response.text();
    }

    // Handle non-OK responses
    if (!response.ok) {
      const errorInfo = {
        status: response.status,
        statusText: response.statusText,
        message: data?.message || data?.error || `HTTP ${response.status}`,
        code: data?.code || 'HTTP_ERROR',
        data: data,
        url: url,
        method: config.method
      };

      console.error('❌ API Error Response:', errorInfo);

      // Auto-handle authentication errors
      if (response.status === 401) {
        console.error('🛑 Authentication failed (401), clearing auth data');
        clearAuthData();
        
        if (!window.location.pathname.includes('/login')) {
          window.location.href = '/login?session=expired';
        }
      }

      if (response.status === 403) {
        console.error('🚫 Access forbidden (403)');
      }

      const error = new Error(errorInfo.message);
      Object.assign(error, errorInfo);
      throw error;
    }

    console.log('✅ API Success:', { endpoint, data: data ? 'received' : 'no data' });
    return data;

  } catch (error) {
    clearTimeout(timeoutId);
    
    // Enhanced error handling
    if (error.name === 'AbortError') {
      error.message = 'Request timeout. Please check your connection.';
      error.code = 'TIMEOUT_ERROR';
    } else if (error.name === 'TypeError') {
      if (error.message.includes('fetch')) {
        error.message = 'Network error. Please check your internet connection.';
        error.code = 'NETWORK_ERROR';
      }
    }

    console.error('💥 API Request Failed:', {
      error: error.message,
      code: error.code,
      endpoint,
      url,
      method: options.method || 'GET'
    });

    throw error;
  }
};

// === HTTP Method Helpers ===
export const get = (endpoint, config = {}) => 
  api(endpoint, { method: 'GET', ...config });

export const post = (endpoint, body, config = {}) => 
  api(endpoint, { method: 'POST', body, ...config });

export const put = (endpoint, body, config = {}) => 
  api(endpoint, { method: 'PUT', body, ...config });

export const patch = (endpoint, body, config = {}) => 
  api(endpoint, { method: 'PATCH', body, ...config });

export const del = (endpoint, config = {}) => 
  api(endpoint, { method: 'DELETE', ...config });

// === Auth API ===
export const authApi = {
  login: async (credentials) => {
    console.log('🔐 Attempting login...', { email: credentials.email });
    
    const response = await post('/users/login', credentials);
    
    // ✅ FIXED: Handle nested response structure from backend
    console.log('📨 Login API Response:', response);
    
    // Extract from nested data object (your backend returns: { success: true, data: { token, user } })
    let userData, token;

    if (response.success && response.data) {
      // New format: nested in data object
      userData = response.data.user;
      token = response.data.token;
    } else if (response.user && response.token) {
      // Old format: flat structure
      userData = response.user;
      token = response.token;
    } else {
      // Direct format
      userData = response;
      token = response.token;
    }

    if (!userData || !token) {
      console.error('❌ Missing user data or token in response:', response);
      throw new Error('Invalid login response: missing user data or token');
    }

    // Normalize user data structure
    const normalizedUser = {
      id: userData.user_id || userData.id,
      user_id: userData.user_id,
      fullname: userData.fullname,
      email: userData.email,
      role: userData.role,
      roles: userData.roles ? 
        (Array.isArray(userData.roles) ? userData.roles : [userData.roles]) : 
        [userData.role || 'user'],
      teacher: userData.teacher,
      is_active: userData.is_active !== undefined ? userData.is_active : true
    };

    // Store auth data
    localStorage.setItem('user', JSON.stringify(normalizedUser));
    localStorage.setItem('token', token);

    console.log('✅ Login successful:', { 
      user: normalizedUser.email, 
      role: normalizedUser.role,
      userId: normalizedUser.user_id,
      tokenLength: token.length 
    });

    return { user: normalizedUser, token };
  },

  logout: () => {
    console.log('🚪 Logging out...');
    clearAuthData();
  },

  getCurrentUser: () => {
    return getStoredUser();
  },

  getToken: () => {
    return getAuthToken();
  },

  validateToken: async () => {
    try {
      const user = getStoredUser();
      const token = getAuthToken();
      
      if (!user || !token) {
        return { valid: false, reason: 'No token or user data' };
      }

      return { valid: true, user };
    } catch (error) {
      return { valid: false, reason: error.message };
    }
  },

  register: async (userData) => {
    console.log('👤 Attempting registration...', { email: userData.email });
    const response = await post('/users/register', userData);
    
    // Handle nested response for registration too
    if (response.success && response.data) {
      return response.data;
    }
    return response;
  }
};

// === Teachers API ===
export const teacherApi = {
  create: (teacher) => post('/teachers', teacher),
  getAll: () => get('/teachers'),
  getById: (id) => get(`/teachers/${id}`),
  update: (id, teacher) => put(`/teachers/${id}`, teacher),
  delete: (id) => del(`/teachers/${id}`),
  getSubjects: (id) => get(`/teachers/${id}/my-subjects`),
  resetPassword: (id) => put(`/teachers/${id}/reset-password`),
  getClasses: (id) => get(`/teachers/${id}/classes`),
};

// === Subjects API ===
export const subjectApi = {
  create: (subject) => post('/subjects', subject),
  getAll: () => get('/subjects'),
  getById: (id) => get(`/subjects/${id}`),
  getStudents: (subjectId) => get(`/subjects/${subjectId}/students`),
  update: (id, subject) => put(`/subjects/${id}`, subject),
  delete: (id) => del(`/subjects/${id}`),
  getTeachers: (subjectId) => get(`/subjects/${subjectId}/teachers`),
};

// === Students API ===
export const studentApi = {
  create: (student) => post('/students', student),
  getAll: () => get('/students'),
  getById: (id) => get(`/students/${id}`),
  update: (id, student) => put(`/students/${id}`, student),
  delete: (id) => del(`/students/${id}`),
  getSubjects: (id) => get(`/students/${id}/subjects`),
  getResults: (id) => get(`/students/${id}/results`),
  getAssignments: (id) => get(`/students/${id}/assignments`),
};

// === Classes API ===
export const classApi = {
  create: (classData) => post('/classes', classData),
  getAll: () => get('/classes'),
  getById: (id) => get(`/classes/${id}`),
  update: (id, classData) => put(`/classes/${id}`, classData),
  delete: (id) => del(`/classes/${id}`),
  getStudents: (id) => get(`/classes/${id}/students`),
  getSubjects: (id) => get(`/classes/${id}/subjects`),
};

// === Streams API ===
export const streamApi = {
  create: (stream) => post('/streams', stream),
  getAll: () => get('/streams'),
  getById: (id) => get(`/streams/${id}`),
  update: (id, stream) => put(`/streams/${id}`, stream),
  delete: (id) => del(`/streams/${id}`),
};

// === Scores API ===
export const scoreApi = {
  create: (score) => post('/scores', score),
  getAll: () => get('/scores'),
  getById: (id) => get(`/scores/${id}`),
  update: (id, score) => put(`/scores/${id}`, score),
  delete: (id) => del(`/scores/${id}`),
  getStudentScores: (studentId) => get(`/scores/student/${studentId}`),
  getSubjectScores: (subjectId) => get(`/scores/subject/${subjectId}`),
};

// === Comments API ===
export const commentApi = {
  create: (comment) => post('/comments', comment),
  getAll: () => get('/comments'),
  getById: (id) => get(`/comments/${id}`),
  update: (id, comment) => put(`/comments/${id}`, comment),
  delete: (id) => del(`/comments/${id}`),
  getStudentComments: (studentId) => get(`/comments/student/${studentId}`),
};

// === Terms API ===
export const termApi = {
  create: (term) => post('/terms', term),
  getAll: () => get('/terms'),
  getById: (id) => get(`/terms/${id}`),
  update: (id, term) => put(`/terms/${id}`, term),
  delete: (id) => del(`/terms/${id}`),
  getCurrent: () => get('/terms/current'),
};

// === Results API ===
export const resultApi = {
  generate: (data) => post('/results/generate', data),
  getAll: () => get('/results'),
  getById: (id) => get(`/results/${id}`),
  getStudentResults: (studentId) => get(`/results/student/${studentId}`),
  getClassResults: (classId) => get(`/results/class/${classId}`),
  publish: (id) => put(`/results/${id}/publish`),
  unpublish: (id) => put(`/results/${id}/unpublish`),
};

// === SMS API ===
export const smsApi = {
  send: (data) => post('/sms/send', data),
  getTemplates: () => get('/sms/templates'),
  createTemplate: (template) => post('/sms/templates', template),
  getHistory: () => get('/sms/history'),
};

// === Assignments API ===
export const assignmentApi = {
  create: (assignment) => post('/assignments', assignment),
  getAll: () => get('/assignments'),
  getById: (id) => get(`/assignments/${id}`),
  update: (id, assignment) => put(`/assignments/${id}`, assignment),
  delete: (id) => del(`/assignments/${id}`),
  getAssignmentStudents: (assignmentId) => get(`/assignments/${assignmentId}/students`),
  submitScore: (assignmentId, studentId, score) => 
    post(`/assignments/${assignmentId}/students/${studentId}/score`, { score }),
  getStudentAssignments: (studentId) => get(`/assignments/student/${studentId}`),
  getTeacherAssignments: (teacherId) => get(`/assignments/teacher/${teacherId}`),
};

// === Users API ===
export const userApi = {
  getAll: () => get('/users'),
  getById: (id) => get(`/users/${id}`),
  update: (id, userData) => put(`/users/${id}`, userData),
  delete: (id) => del(`/users/${id}`),
  changePassword: (id, passwordData) => put(`/users/${id}/password`, passwordData),
  resetPassword: (email) => post('/users/reset-password', { email }),
};

// Export API base URL for external use
export { API_BASE_URL };
