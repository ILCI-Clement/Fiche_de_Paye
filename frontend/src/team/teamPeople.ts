import { accountStatus } from '../admin/accountStatus.ts'
import type { Personnel } from './api'

/** Business function and login privileges are independent: managers are people too. */
export function teamPeople(people: Personnel[]): Personnel[] {
  return people.filter(person => accountStatus(person) === 'active' &&
    (person.role_tags?.length ? person.role_tags.some(role => role === 'Employe' || role === 'Responsable') :
      person.role === 'Employe' || person.role === 'Responsable'))
}
