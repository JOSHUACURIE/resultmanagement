
export const ROLES = {
  admin: 'ADMIN',
  dos: 'DOS',
  teacher: 'TEACHER',
  principal: 'PRINCIPAL'
};


export const getRoleDisplayName = (role) => {
  const normalized = typeof role === 'string' ? role.toUpperCase() : String(role).toUpperCase();
  const displayNames = {
    [ROLES.admin]: 'Administrator',
    [ROLES.dos]: 'Director of Studies',
    [ROLES.teacher]: 'Teacher',
    [ROLES.principal]: 'Principal'
  };
  return displayNames[normalized] || role;
};

export const isValidRole = (role) => {
  if (!role) return false;
  const normalized = typeof role === 'string' ? role.toUpperCase() : String(role).toUpperCase();
  return Object.values(ROLES).includes(normalized);
};