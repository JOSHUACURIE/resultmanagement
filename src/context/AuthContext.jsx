import React, { createContext, useState, useContext, useEffect } from 'react';

const AuthContext = createContext();

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
};

export const AuthProvider = ({ children }) => {
  const [user, setUser] = useState(null);
  const [token, setToken] = useState(null);
  const [loading, setLoading] = useState(true);

  // Validate user data structure
  const validateUserData = (userData) => {
    if (!userData) {
      throw new Error('Invalid user data: user data is required');
    }

    const requiredFields = ['user_id', 'email', 'role'];
    const missingFields = requiredFields.filter(field => !userData[field]);

    if (missingFields.length > 0) {
      throw new Error(`Invalid user data: missing required fields (${missingFields.join(', ')})`);
    }

    // Validate role - convert to lowercase for consistency
    const validRoles = ['dos', 'principal', 'teacher'];
    const userRole = userData.role.toLowerCase();
    if (!validRoles.includes(userRole)) {
      throw new Error(`Invalid user data: invalid role '${userData.role}'`);
    }

    // Return normalized user data with lowercase role
    return {
      ...userData,
      role: userRole
    };
  };

  // Login function
  const login = (userData, authToken) => {
    try {
      // Validate user data before setting
      const validatedUserData = validateUserData(userData);
      
      setUser(validatedUserData);
      setToken(authToken);
      
      // Store in localStorage
      localStorage.setItem('user', JSON.stringify(validatedUserData));
      localStorage.setItem('token', authToken);
      
      console.log('✅ Login successful:', { 
        user: validatedUserData, 
        hasToken: !!authToken 
      });
    } catch (error) {
      console.error('Login validation error:', error);
      throw error;
    }
  };

  // Logout function
  const logout = () => {
    console.log('🚪 Logging out user');
    setUser(null);
    setToken(null);
    localStorage.removeItem('user');
    localStorage.removeItem('token');
  };

  // Check if user is authenticated
  const isAuthenticated = () => {
    return !!user && !!token;
  };

  // Check if user has specific role (case-insensitive)
  const hasRole = (role) => {
    if (!user || !role) return false;
    return user.role.toLowerCase() === role.toLowerCase();
  };

  // Get user role (always lowercase for consistency)
  const getUserRole = () => {
    return user?.role?.toLowerCase() || null;
  };

  // Initialize auth state from localStorage
  useEffect(() => {
    const initializeAuth = async () => {
      try {
        const storedUser = localStorage.getItem('user');
        const storedToken = localStorage.getItem('token');

        console.log('🔄 Initializing auth from localStorage:', {
          hasStoredUser: !!storedUser,
          hasStoredToken: !!storedToken
        });

        if (storedUser && storedToken) {
          const userData = JSON.parse(storedUser);
          // Validate stored user data
          const validatedUserData = validateUserData(userData);
          setUser(validatedUserData);
          setToken(storedToken);
          
          console.log('✅ Auth initialized successfully');
        } else {
          console.log('ℹ️ No stored auth data found');
        }
      } catch (error) {
        console.error('❌ Error initializing auth:', error);
        // Clear invalid stored data
        localStorage.removeItem('user');
        localStorage.removeItem('token');
      } finally {
        setLoading(false);
      }
    };

    initializeAuth();
  }, []);

  const value = {
    user,
    token,
    login,
    logout,
    isAuthenticated,
    hasRole,
    getUserRole, // Added this new function
    loading
  };

  return (
    <AuthContext.Provider value={value}>
      {children}
    </AuthContext.Provider>
  );
};

export default AuthContext;