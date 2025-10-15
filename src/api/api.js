// src/api/api.js
const API_BASE_URL = 'https://result-6.onrender.com/';

// === Helper: Get stored user ===
const getStoredUser = () => {
  const raw = localStorage.getItem('user');
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    console.error('Failed to parse user from localStorage', e);
    return null;
  }
};
export const assignmentApi = {
  getAssignmentStudents: (assignmentId) => {
    return get(`/assignments/${assignmentId}/students`);
  },
};
// === Helper: Get stored token ===
const getAuthToken = () => {
  // Get token directly from localStorage (stored separately by AuthContext)
  const token = localStorage.getItem('token');
  return token || null;
};

// === Core API function ===
const api = async (endpoint, options = {}) => {
  const url = `${API_BASE_URL}${endpoint}`;

  // Debug logging
  console.log('🔄 API Request:', {
    url,
    method: options.method || 'GET',
    endpoint,
    hasBody: !!options.body
  });

  const headers = {
    'Content-Type': 'application/json',
    ...options.headers,
  };

  const token = getAuthToken();
  if (token) {
    headers['Authorization'] = `Bearer ${token}`;
    console.log('🔐 Using auth token:', token ? '***' + token.slice(-10) : 'none');
  } else {
    console.warn('⚠️ No auth token found');
  }

  const config = {
    method: options.method || 'GET',
    headers,
  };

  // Add body if it exists (skip for GET/DELETE or FormData)
  if (options.body && !['GET', 'DELETE'].includes(options.method || 'GET')) {
    if (typeof options.body === 'object' && !(options.body instanceof FormData)) {
      config.body = JSON.stringify(options.body);
    } else {
      config.body = options.body;
    }
  }

  try {
    const response = await fetch(url, config);

    console.log('📡 API Response:', {
      status: response.status,
      ok: response.ok,
      url: response.url,
      contentType: response.headers.get('content-type')
    });

    const contentType = response.headers.get('content-type');
    let data;
    if (contentType && contentType.includes('application/json')) {
      data = await response.json().catch(() => ({}));
    } else {
      data = await response.text();
    }

    console.log('📦 Response data:', data);

    if (!response.ok) {
      const errorMessage = data?.message || data?.error || data || `HTTP Error: ${response.status}`;
      const error = new Error(errorMessage);
      error.status = response.status;
      error.code = data?.code || 'UNKNOWN_ERROR';
      error.data = data;
      
      // Auto-logout on 401 Unauthorized
      if (response.status === 401) {
        console.error('🛑 Authentication failed, clearing tokens');
        localStorage.removeItem('user');
        localStorage.removeItem('token');
        window.location.href = '/login';
      }
      
      throw error;
    }

    return data;
  } catch (error) {
    console.error('❌ API Error:', error);
    if (error.name === 'TypeError' && /fetch/.test(error.message)) {
      error.message = 'Network error. Please check your connection.';
      error.code = 'NETWORK_ERROR';
    }
    throw error;
  }
};

// === Generic helpers ===
export const get = (endpoint, config = {}) => api(endpoint, { method: 'GET', ...config });
export const post = (endpoint, body, config = {}) => api(endpoint, { method: 'POST', body, ...config });
export const put = (endpoint, body, config = {}) => api(endpoint, { method: 'PUT', body, ...config });
export const del = (endpoint, config = {}) => api(endpoint, { method: 'DELETE', ...config });

// === Auth API ===
export const authApi = {
  login: async (credentials) => {
    const response = await post('/users/login', credentials);

    // Normalize roles to always be an array
    if (response && response.roles && !Array.isArray(response.roles)) {
      response.roles = [response.roles];
    }

    // Save user and token separately (matching AuthContext structure)
    if (response.user && response.token) {
      localStorage.setItem('user', JSON.stringify(response.user));
      localStorage.setItem('token', response.token);
    } else {
      // Backward compatibility if response structure is different
      localStorage.setItem('user', JSON.stringify(response));
      if (response.token) {
        localStorage.setItem('token', response.token);
      }
    }

    return response;
  },
  logout: () => {
    localStorage.removeItem('user');
    localStorage.removeItem('token');
  },
  getCurrentUser: () => {
    return getStoredUser();
  },
  getToken: () => {
    return getAuthToken();
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
};

// === Subjects API ===
export const subjectApi = {
  create: (subject) => post('/subjects', subject),
  getAll: () => get('/subjects'),
  getStudents: (subjectId) => get(`/subjects/${subjectId}/students`),
  update: (id, subject) => put(`/subjects/${id}`, subject),
  delete: (id) => del(`/subjects/${id}`),
};

// === Students API ===
export const studentApi = {
  create: (student) => post('/students', student),
  getAll: () => get('/students'),
  getById: (id) => get(`/students/${id}`),
  update: (id, student) => put(`/students/${id}`, student),
  delete: (id) => del(`/students/${id}`),
};


