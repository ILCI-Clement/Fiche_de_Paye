import { useQuery } from '@tanstack/react-query'
import { Alert, Box, CircularProgress } from '@mui/material'
import { adminApi, type OrganizationInput } from '../admin/api'
import { OrganizationFields } from '../admin/OrganizationFields'

export function RegistrationAccountFields({ token, value, onChange, busy, allowSuperAdmin }: {
  token: string; value: OrganizationInput; onChange: (value: OrganizationInput) => void;
  busy: boolean; allowSuperAdmin: boolean;
}) {
  const users = useQuery({ queryKey: ['registration-account-users'], queryFn: ({ signal }) => adminApi.users(token, signal), gcTime: 0, retry: false })
  const groups = useQuery({ queryKey: ['registration-account-groups'], queryFn: ({ signal }) => adminApi.groups(token, signal), gcTime: 0, retry: false })
  if (users.isPending || groups.isPending) return <CircularProgress size={22} />
  if (users.error || groups.error) return <Alert severity="error">Impossible de charger les départements et responsables. Fermez puis réessayez.</Alert>
  return <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
    <OrganizationFields value={value} onChange={onChange} users={users.data.users} groups={groups.data.groups}
      disabled={busy} allowSuperAdmin={allowSuperAdmin} hidePersonnelType />
  </Box>
}
