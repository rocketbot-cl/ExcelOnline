from http import client
from re import A
import requests
import json
import os
import msal
import webbrowser
class ExcelOnlineService:
    
    def __init__(self, *, client_id, client_secret, path_user, path_credentials):
        self.client_id = client_id
        self.client_secret = client_secret
        self.authority_url = 'https://login.microsoftonline.com/consumers/'
        self.scopes = ['User.Read', 'Files.ReadWrite.All']
        self.access_token = None
        self.refresh_token = None
        self.path_user = path_user
        self.path_credentials = path_credentials
        self.base_url = 'https://graph.microsoft.com/v1.0/'
    def get_code(self):
        try:
            client_instance = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority_url
            )
            authorization_request_url = client_instance.get_authorization_request_url(self.scopes)
           
            webbrowser.open(authorization_request_url, new=True)
            
            with open(self.path_user, 'w') as userfile:
                user_info = {'client_id': self.client_id, 'client_secret': self.client_secret}
                json.dump(user_info, userfile)
            return client_instance
        
        except Exception as e:
            error_info ={
                'error': str(e),
                'error_description': 'Error in get_code',
            }
            print(error_info)
            raise e
    
    
    def get_new_token(self, client_instance, auth_code):
        """ Get the access_token or refresh_token.

        The token that is obtained depends if it is the first time that the user is authenticated or not.

        Parameters
        ----------
        auth_code : dict
            Contains code or refresh_token
        grant_type : str
            Type of grant_type, it could be code or refresh_token

        Returns
        -------
        dict
            a json with the credentials
        """
        
        try:           
            access_token = client_instance.acquire_token_by_authorization_code(
                code = auth_code,
                scopes = self.scopes
            )
            print(access_token)
            print(type(access_token))
            json_response = access_token
            self.access_token = json_response['access_token']
            self.refresh_token = json_response['refresh_token']
            return json_response
        
        except Exception as e:
            error_info ={
                'error': str(e),
                'error_description': 'Error in get_token',
            }
            print(error_info)
            raise e

    def get_old_token(self, refresh_token):
        """ Build the request.

        It depends if it is access_token or refresh_token.

        Parameters
        ----------
        auth_code : dict
            Contains code or refresh_token
        grant_type : str
            Type of grant_type, it could be code or refresh_token

        Returns
        -------
        string, dict
            a formed url and a dict with parameters

        """
        try:
            client_instance = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=self.authority_url
            )
            access_token = client_instance.acquire_token_by_refresh_token(
                refresh_token = refresh_token,
                scopes = self.scopes
            )
            json_response = access_token
            self.access_token = json_response['access_token']
            self.refresh_token = json_response['refresh_token']
            return json_response
        except Exception as e:
            error_info ={
                'error': str(e),
                'error_description': 'Error in get_token',
            }
            print(error_info)
            raise e

    def create_tokens_file(self, credentials):
        """ Create a json with credentials.

        Create a json with credentials.

        Parameters
        ----------
        credentials : dict
            Contains the credentials

        """
        try:
            with open(self.path_credentials, 'w') as credfile:
                json.dump(credentials, credfile)
            return True
        except Exception as e:
            print(e)
            raise e

    def get_xlsx_files(self):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = self.base_url + "me/drive/root/search(q='.xlsx')?select=name,id,webUrl"
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def get_worksheets(self, workbook_id):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets".format(workbook_id=workbook_id)
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def create_workbook(self, name):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = self.base_url + f"/me/drive/root:/{name}.xlsx:/content".format(name=name)
        response = requests.put(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def add_new_worksheet(self, workbook_id, sheet_name):
        headers = {
            # 'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + self.access_token,
            # 'workbook-session-id': session_id
        }
        data = {
            "name": sheet_name
        }
        url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets/".format(
            workbook_id=workbook_id)
        response = requests.post(url, json=data, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    # def create_session(self, id):
    #     headers = {
    #         'Authorization': 'Bearer ' + self.access_token
    #     }
    #     session_params = {
    #         "persistChanges": True
    #     }
    #     url = "https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/createSession".format(id=id)
    #     response = requests.post(url, json=session_params, headers=headers)
    #     print("Response: ", response.text)
    #     json_response = json.loads(response.text)
    #     print("Json: ", json_response)
    #     return json_response

    def get_cell(self, workbook_id, sheet_name, range_cell): # session_id
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
            # 'workbook-session-id': session_id
        }
        url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')".format(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range_cell=range_cell
        )
        
        print(range_cell)
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def update_range(self, workbook_id, sheet_name, range_cell, value_cell): # session_id
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
            # 'workbook-session-id': session_id
        }
        cells = self.get_cell(workbook_id, sheet_name, range_cell)['values'] # session_id
        new_values = []
        for value in cells:
            replaced_cells = [value_cell for x in value]
            new_values.append(replaced_cells)
        data = {
            "values": new_values
        }
        url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')".format(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range_cell=range_cell
        )
        response = requests.patch(url, json=data, headers=headers)
        print("Response: ", response.text)
        json_response = json.loads(response.text)
        print("Json: ", json_response)
        return json_response