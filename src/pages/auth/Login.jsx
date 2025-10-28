import React, { useState } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { authApi } from '../../api/api';
import Button from '../../components/ui/Button';
import Input from '../../components/ui/Input';
import { useAuth } from '../../context/AuthContext';
import './Login.css';

const Login = () => {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const navigate = useNavigate();
  const location = useLocation();
  const { login } = useAuth();

  // Valid roles in the system
  const VALID_ROLES = ['dos', 'principal', 'teacher'];

  // Normalize role to lowercase and validate
  const normalizeRole = (role) => {
    if (!role) return null;
    
    const normalizedRole = role.toLowerCase().trim();
    return VALID_ROLES.includes(normalizedRole) ? normalizedRole : null;
  };

  // Extract and validate user role from response
  const extractUserRole = (userData) => {
    // Priority 1: Check role field (new schema)
    if (userData.role) {
      return normalizeRole(userData.role);
    }

    // Priority 2: Check roles array (legacy schema)
    if (userData.roles) {
      if (Array.isArray(userData.roles) && userData.roles.length > 0) {
        return normalizeRole(userData.roles[0]);
      }
      if (typeof userData.roles === 'string') {
        return normalizeRole(userData.roles);
      }
    }

    // Priority 3: Check teacher profile
    if (userData.teacher && userData.teacher.teacher_id) {
      return 'teacher';
    }

    return null;
  };

  // Validate user data structure
  const validateUserData = (userData) => {
    if (!userData) {
      throw new Error('No user data received');
    }

    // Check for required fields
    const requiredFields = ['user_id', 'email'];
    const missingFields = requiredFields.filter(field => !userData[field]);

    if (missingFields.length > 0) {
      throw new Error(`Invalid user data: missing required fields (${missingFields.join(', ')})`);
    }

    // Validate role
    const userRole = extractUserRole(userData);
    if (!userRole) {
      throw new Error('No valid role assigned to this account');
    }

    return {
      user_id: userData.user_id,
      email: userData.email,
      role: userRole,
      fullname: userData.fullname || userData.full_name || '',
      teacher_id: userData.teacher_id || (userData.teacher ? userData.teacher.teacher_id : null)
    };
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    // Basic validation
    if (!email.trim() || !password.trim()) {
      setError('Please enter both email and password');
      setLoading(false);
      return;
    }

    try {
      // Call login API - it already returns { user, token }
      const result = await authApi.login({ 
        email: email.toLowerCase().trim(), 
        password: password.trim() 
      });

      console.log('Login API Result:', result); // Debug log

      // ✅ FIXED: Use the result directly as returned by authApi.login()
      const { user: userData, token } = result;

      if (!userData || !token) {
        throw new Error('Invalid response from server: missing user data or token');
      }

      // Validate and normalize user data
      const validatedUserData = validateUserData(userData);

      // ✅ Save user & token in context
      login(validatedUserData, token);

      // Redirect based on role
      const from = location.state?.from?.pathname || getDashboardPath(validatedUserData.role);
      navigate(from, { replace: true });

    } catch (err) {
      console.error('Login error:', err);
      
      // ✅ FIXED: Create new error message instead of modifying error object
      let errorMessage = 'Login failed. Please check your credentials and try again.';
      
      if (err.response?.data?.message) {
        errorMessage = err.response.data.message;
      } else if (err.message) {
        errorMessage = err.message;
      } else if (err.includes?.('Network Error')) {
        errorMessage = 'Network error. Please check your connection and try again.';
      }
      
      setError(errorMessage);
    } finally {
      setLoading(false);
    }
  };

  const getDashboardPath = (userRole) => {
    switch (userRole) {
      case 'dos':
        return '/dos/dashboard';
      case 'principal':
        return '/principal/dashboard';
      case 'teacher':
        return '/teacher/submit-scores';
      default:
        return '/login';
    }
  };

  return (
    <div className="login-container">
      <div className="login-card">
        <div className="login-header">
          <h1 className="login-title">School Result Management</h1>
          <p className="login-subtitle">Sign in to your account</p>
        </div>

        <form onSubmit={handleSubmit} className="login-form">
          {error && (
            <div className="login-error">
              <strong>Error:</strong> {error}
            </div>
          )}

          <Input
            label="Email Address"
            id="email"
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
            placeholder="e.g., name@school.edu"
            autoComplete="email"
          />

          <Input
            label="Password"
            id="password"
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
            placeholder="••••••••"
            autoComplete="current-password"
          />

          <Button
            type="submit"
            variant="primary"
            size="lg"
            disabled={loading || !email.trim() || !password.trim()}
            className="login-submit-btn"
            fullWidth
          >
            {loading ? 'Signing in...' : 'Sign In'}
          </Button>
        </form>

        <div className="login-footer">
          <p>
            Having trouble? Contact{' '}
            <a href="mailto:support@school.edu" className="support-link">
              support@school.edu
            </a>
          </p>
        </div>
      </div>
    </div>
  );
};

export default Login;
