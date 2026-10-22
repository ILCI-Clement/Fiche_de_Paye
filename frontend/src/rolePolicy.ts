type RoleSubject = { role_tags: readonly string[] }

export const isSuperAdmin = (subject: RoleSubject) => subject.role_tags.includes('SuperAdmin')
export const isAdmin = (subject: RoleSubject) => isSuperAdmin(subject) || subject.role_tags.includes('Admin')
export const canManageTeam = (subject: RoleSubject) => isAdmin(subject) || subject.role_tags.includes('Responsable')
export const roleLabel = (role: string) => ({ SuperAdmin: 'Super Admin', Admin: 'Admin', Responsable: 'Responsable', Employe: 'Employé' })[role] || role
export const accessRoleLabel = (roles: readonly string[]) => roles.includes('SuperAdmin') ? roleLabel('SuperAdmin') : roles.includes('Admin') ? roleLabel('Admin') : 'Standard'
export const displayFunctionRoles = (roles: readonly string[]) => ['Responsable', 'Employe'].filter(role => roles.includes(role)).map(roleLabel).join(' · ')
export const displayRoles = (roles: readonly string[]) => [...new Set(['SuperAdmin', 'Admin', 'Responsable', 'Employe', ...roles])].filter(role => roles.includes(role) && (role !== 'Admin' || !roles.includes('SuperAdmin'))).map(roleLabel).join(' · ')
