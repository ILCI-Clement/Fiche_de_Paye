"""Manager attendance rights without database writes or email dispatch."""
import copy
import unittest
from datetime import date
from unittest.mock import patch
import test_backend_helpers  # supplies isolated environment defaults
from backend import main
from backend.attendance_permissions import validate_manager_save
from contract_periods import bind_contract
from fastapi.testclient import TestClient

MANAGER = {'username':'manager','name':'manager','role':'Responsable','role_tags':['Responsable'],'managed_group_ids':[1]}
EMPLOYEE = {'username':'employee','email':'employee@example.test','role_tags':['Employe'],'group_ids':[1],'manager_id':'manager'}
CONTRACT = {'id':1,'employee_type':'salarie','start_date':date(2026,9,1),'end_date':date(2026,9,30),'is_cdi':False}


class Cursor:
    def execute(self, *args): pass
    def fetchone(self): return CONTRACT


class ManagerPermissionTests(unittest.TestCase):
    def check_save(self, previous, incoming, actor=MANAGER, visible=True):
        return validate_manager_save(Cursor(),actor,previous,incoming,lambda *args:EMPLOYEE,lambda *args:visible)

    def test_manager_can_change_attendance(self):
        original={'id':'sheet','account_username':'employee','nom':'Employee','calendar_overrides':{}}
        changed=copy.deepcopy(original); changed['calendar_overrides']={'2026-09-02':{'matin':{'status':'Absence'}}}
        self.check_save({'employes_data':[original]},{'employes_data':[changed]})

    def test_manager_cannot_delete_sheet(self):
        with self.assertRaises(main.HTTPException) as e:
            self.check_save({'employes_data':[{'id':'sheet'}]},{'employes_data':[]})
        self.assertEqual(e.exception.status_code,403)

    def test_manager_cannot_modify_identity(self):
        original={'id':'sheet','account_username':'employee','nom':'Employee'}
        with self.assertRaises(main.HTTPException):
            self.check_save({'employes_data':[original]},{'employes_data':[{**original,'nom':'Other'}]})

    def test_manager_cannot_edit_outside_team(self):
        with self.assertRaises(main.HTTPException):
            self.check_save({}, {'employes_data':[{'id':'sheet','account_username':'other'}]},visible=False)

    def test_manager_can_create_registered_contract_sheet(self):
        fiche=main.serialize_fiche_config({'employes_data':[bind_contract({'account_username':'employee'},CONTRACT,9,2026)]})
        import json
        self.check_save({},json.loads(fiche))

    def test_admin_cannot_delete_sheet(self):
        with self.assertRaises(main.HTTPException) as e:
            self.check_save({'employes_data':[{'id':'sheet'}]}, {'employes_data':[]}, actor={'role_tags':['Admin']})
        self.assertEqual(e.exception.status_code, 403)

    def test_manager_account_routes_forbidden(self):
        main.app.dependency_overrides[main.get_current_user]=lambda:MANAGER
        try:
            client=TestClient(main.app)
            for method,path,payload in [('POST','/create-user',{}),('PUT','/users/employee/organization',{}),('DELETE','/delete-user/employee',None),('PATCH','/annual-interviews/employee/2026',{})]:
                with self.subTest(path=path):
                    result=client.request(method,path,json=payload)
                    self.assertEqual(result.status_code,403,result.text)
        finally:
            main.app.dependency_overrides.clear()


    def test_manager_cannot_create_contract(self):
        main.app.dependency_overrides[main.get_current_user]=lambda:MANAGER
        try:
            result=TestClient(main.app).post('/users/employee/contracts',json={'employee_type':'salarie','start_date':'2026-09-01','end_date':'2026-09-30','is_cdi':False})
            self.assertEqual(result.status_code,403)
        finally:
            main.app.dependency_overrides.clear()
