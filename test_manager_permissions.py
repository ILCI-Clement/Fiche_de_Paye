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
from streamlit.testing.v1 import AppTest

MANAGER = {'username':'manager','name':'manager','role':'Responsable','role_tags':['Responsable'],'managed_group_ids':[1]}
EMPLOYEE = {'username':'employee','email':'employee@example.test','role_tags':['Employe'],'group_ids':[1],'manager_id':'manager'}
CONTRACT = {'id':1,'employee_type':'salarie','start_date':date(2026,9,1),'end_date':date(2026,9,30),'is_cdi':False}


class Cursor:
    def execute(self, *args): pass
    def fetchone(self): return CONTRACT


class Response:
    status_code=200
    def __init__(self,data): self.data=data
    def json(self): return self.data


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

    def test_manager_people_page_has_no_create_form(self):
        def response(method,url,**kwargs):
            if url.endswith('/organization-directory'):
                return Response({'groups':[{'id':1,'name':'Team','is_active':True}], 'users':[EMPLOYEE]})
            return Response({'groups':[{'id':1,'name':'Team','is_active':True}]} if url.endswith('/groups') else {'users':[EMPLOYEE]})
        with patch('streamlit.page_link'), patch('team_api.safe_api_request',side_effect=response), patch('api_client.safe_api_request',side_effect=response), patch('team_api.api_url',return_value='https://api.test'), patch('team_api.authenticated_headers',return_value={}):
            app=AppTest.from_file('pages/People.py')
            app.session_state['user']=MANAGER
            app.run()
            self.assertFalse(app.exception)
            self.assertEqual(app.title[0].value,'Mon équipe')
            self.assertFalse(app.text_input)

    def test_manager_fiches_hides_delete_and_manual_add(self):
        def response(method,url,**kwargs):
            return Response({'users':[EMPLOYEE]})
        import types
        docx=types.ModuleType('DocxGen'); docx.generer_docx_stagiaire=lambda *args:None
        excel=types.ModuleType('ExcelGen'); excel.remplir_fiche_paie=lambda *args:None
        with patch.dict('sys.modules',{'DocxGen':docx,'ExcelGen':excel}), patch('api_client.safe_api_request',side_effect=response), patch('api_client.api_url',return_value='https://api.test'), patch('api_client.authenticated_headers',return_value={}):
            app=AppTest.from_file('pages/Fiches.py')
            app.session_state['user']=MANAGER
            fiche=bind_contract({'account_username':'employee','nom':'Employee','responsable':'Manager', 'email_employe':'employee@example.test','vacances':[],'absences':[],'arret':[], 'planning_detail':{}}, CONTRACT,9,2026)
            app.session_state['user_data']={'manager':{'mois':9,'annee':2026,'employes_data':[fiche]}}
            app.run()
            self.assertFalse(app.exception)
            labels=[b.label for b in app.button]
            self.assertNotIn('Ajouter un employé / stagiaire',labels)
            self.assertNotIn('Supprimer cette fiche',labels)
            self.assertIn('Ouvrir / créer sa fiche',labels)
            self.assertTrue(all(widget.disabled for widget in app.text_input))
            self.assertTrue(app.radio[0].disabled)

    def test_sources_compile_including_admin(self):
        from pathlib import Path
        for path in [Path('pages/Admin.py'),Path('pages/Fiches.py'),Path('backend/main.py')]:
            compile(path.read_text(encoding='utf-8'),str(path),'exec')

    def test_manager_can_fill_but_not_edit_personnel(self):
        from access_control import can_edit_employee,can_fill_attendance
        self.assertFalse(can_edit_employee(MANAGER,EMPLOYEE))
        self.assertTrue(can_fill_attendance(MANAGER,EMPLOYEE))

    def test_manager_cannot_create_contract(self):
        main.app.dependency_overrides[main.get_current_user]=lambda:MANAGER
        try:
            result=TestClient(main.app).post('/users/employee/contracts',json={'employee_type':'salarie','start_date':'2026-09-01','end_date':'2026-09-30','is_cdi':False})
            self.assertEqual(result.status_code,403)
        finally:
            main.app.dependency_overrides.clear()

    def test_renaming_second_sheet_keeps_active_sheet(self):
        import types
        docx=types.ModuleType('DocxGen'); docx.generer_docx_stagiaire=lambda *args:None
        excel=types.ModuleType('ExcelGen'); excel.remplir_fiche_paie=lambda *args:None
        base={'type':'Salarié','responsable':'Manager','vacances':[],'absences':[],'arret':[], 'planning_detail':{},'ddc':None,'fdc':None,'cdi':False}
        first={**copy.deepcopy(base),'id':'first','nom':'First'}
        second={**copy.deepcopy(base),'id':'second','nom':'Second'}
        with patch.dict('sys.modules',{'DocxGen':docx,'ExcelGen':excel}), patch('api_client.safe_api_request',return_value=Response({'users':[]})), patch('api_client.api_url',return_value='https://api.test'), patch('api_client.authenticated_headers',return_value={}):
            app=AppTest.from_file('pages/Fiches.py')
            app.session_state['user']={'name':'admin','role':'Admin','role_tags':['Admin']}
            app.session_state['user_data']={'admin':{'mois':9,'annee':2026,'employes_data':[first,second]}}
            app.session_state['fiche_active_id_admin_2026_9']='second'
            app.run()
            self.assertFalse(app.exception)
            name=next(w for w in app.text_input if w.label=='NOM Prénom (Employé)')
            self.assertEqual(name.value,'Second')
            name.set_value('Renamed second').run()
            app.run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state['fiche_active_id_admin_2026_9'],'second')
            self.assertEqual(app.session_state['fiche_active_tab'],'Renamed second')
            self.assertEqual(next(w for w in app.text_input if w.label=='NOM Prénom (Employé)').value,'Renamed second')
