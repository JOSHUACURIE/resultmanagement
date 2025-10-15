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
    // Replace with your actual Render backend URL
    return 'https://your-backend-app.onrender.com/api';
  }
  
  // Fallback - UPDATE THIS WITH YOUR ACTUAL RENDER URL
  return 'https://your-backend-app.onrender.com/api';
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
    credentials: 'include', // Important for CORS with credentials
  };

  // Handle request body
  if (options.body && !['GET', 'HEAD'].includes(config.method)) {
    if (options.body instanceof FormData) {
      // Remove Content-Type for FormData to let browser set it
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
        
        // Only redirect if not already on login page
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
    
    // Normalize response structure
    const userData = response.user || response;
    const token = response.token;

    if (!userData || !token) {
      throw new Error('Invalid login response: missing user data or token');
    }

    // Normalize roles to always be an array
    if (userData.roles && !Array.isArray(userData.roles)) {
      userData.roles = [userData.roles];
    }

    // Store auth data
    localStorage.setItem('user', JSON.stringify(userData));
    localStorage.setItem('token', token);

    console.log('✅ Login successful:', { 
      user: userData.email, 
      roles: userData.roles,
      tokenLength: token.length 
    });

    return { user: userData, token };
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

      // You might want to call a validation endpoint here
      // const response = await get('/users/validate');
      // return { valid: true, user: response.user };
      
      return { valid: true, user };
    } catch (error) {
      return { valid: false, reason: error.message };
    }
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

// === Assignments API ===
export const assignmentApi = {
  getAssignmentStudents: (assignmentId) => get(`/assignments/${assignmentId}/students`),
  create: (assignment) => post('/assignments', assignment),
  getAll: () => get('/assignments'),
  getById: (id) => get(`/assignments/${id}`),
  update: (id, assignment) => put(`/assignments/${id}`, assignment),
  delete: (id) => del(`/assignments/${id}`),
};

// Export API base URL for external use
export { API_BASE_URL };
