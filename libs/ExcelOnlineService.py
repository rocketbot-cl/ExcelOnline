import requests
import json
import os


class ExcelOnlineService:
    def __init__(self, *, client_id, client_secret, tenant, redirect_uri, path_credentials):
        self.client_id = client_id
        self.client_secret = client_secret
        self.scope = 'files.readwrite.all offline_access'
        self.redirect_uri = redirect_uri
        self.tenant = tenant
        self.access_token = None
        self.refresh_token = None
        self.path_credentials = path_credentials

    def get_token(self, auth_code, grant_type):
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

        url, params = self.build_request(auth_code, grant_type)
        try:
            response = requests.post(url, data=params)
            json_response = json.loads(response.text)
            self.access_token = json_response['access_token']
            new_refresh_token = json_response['refresh_token']
            self.refresh_token = new_refresh_token
            return json_response
        except Exception as e:
            error_info ={
                'error': str(e),
                'error_description': 'Error in get_token',
                'status_code': response.status_code,
                'response': response.text
            }
            print(error_info)
            raise e

    def build_request(self, auth_code, grant_type):
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

        params = {
            'client_id': self.client_id,
            'scope': self.scope,
            'redirect_uri': self.redirect_uri,
            'grant_type': grant_type,
            'client_secret': self.client_secret,
        }
        params.update(auth_code)
        url = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'.format(tenant=self.tenant)
        return url, params

    def create_tokens_file(self, credentials):
        """ Create a json with credentials.

        Create a json with credentials.

        Parameters
        ----------
        credentials : dict
            Contains the credentials

        """
        try:
            with open(self.path_credentials, 'w') as outfile:
                json.dump(credentials, outfile)
            return True
        except Exception as e:
            print(e)
            raise e

    def get_xlsx_files(self):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='.xlsx')?select=name,id,webUrl"
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def get_worksheets(self, workbook_id):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = "https://graph.microsoft.com/v1.0/me/drive/items/{workbook_id}/workbook/worksheets".format(workbook_id=workbook_id)
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def create_workbook(self, name):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = "https://graph.microsoft.com/v1.0/me/drive/root:/{name}.xlsx:/content".format(name=name)
        response = requests.put(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def add_new_worksheet(self, session_id, workbook_id, sheet_name):
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        data = {
            "name": sheet_name
        }
        url = "https://graph.microsoft.com/v1.0/me/drive/items/{workbook_id}/workbook/worksheets/".format(
            workbook_id=workbook_id)
        response = requests.post(url, json=data, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def create_session(self, id):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        session_params = {
            "persistChanges": True
        }
        url = "https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/createSession".format(id=id)
        response = requests.post(url, json=session_params, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def get_cell(self, session_id, workbook_id, sheet_name, range_cell):
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        url = "https://graph.microsoft.com/v1.0/me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range}')".format(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range=range_cell
        )
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response

    def update_range(self, session_id, workbook_id, sheet_name, range_cell, value_cell):
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        cells = self.get_cell(session_id, workbook_id, sheet_name, range_cell)['values']
        new_values = []
        for value in cells:
            replaced_cells = [value_cell for x in value]
            new_values.append(replaced_cells)
        data = {
            "values": new_values
        }
        url = "https://graph.microsoft.com/v1.0/me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')".format(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range_cell=range_cell
        )
        response = requests.patch(url, json=data, headers=headers)
        json_response = json.loads(response.text)
        return json_response
