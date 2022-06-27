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
        """ Get the authorization code.

        The code is needed to obtain the access credentials.

        Returns
        -------
        dict, client
            a json with client data and a client instance
        """
        
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
        """ Get the access_token.

        The token obtained is the one of the first time that the user is authenticated.

        Parameters
        ----------
        client_instance : client instance 
            Initialized with get_code()
        auth_code : str
            Authorization Code for the client instance

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
        """ Get the access_token using the refresh_token.

        The token that is obtained is the one after the user has been authenticated for the first time.

        Parameters
        ----------
        refresh_token : str
            
        Returns
        -------
        dict
            a json with the credentials
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
        """ Get the '.xlsx' files in the directory.
            
        Returns
        -------
        list
            a list of dictionaries with name and ID of the xlsx files
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = self.base_url + "me/drive/root/search(q='.xlsx')?select=name,id,webUrl"
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        clean_data = []
        for xlsx in json_response['value']:
            for k in ['@odata.type','webUrl']:
                xlsx.pop(k)
            clean_data.append(xlsx)
        return clean_data

    def get_worksheets(self, workbook_id):
        """ Get the worksheets of a workbook.

        Parameters
        ----------
        workboook_id : str
            
        Returns
        -------
        list
            a list with the worksheets names
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets".format(workbook_id=workbook_id)
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        sheets = []
        for sheet in json_response['value']:
            sheets.append(sheet['name'])
        return sheets

    def create_workbook(self, name):
        """ Creates a new workbook.

        Parameters
        ----------
        name : str
            
        Returns
        -------
        str
            the ID of the new workbook
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        url = self.base_url + f"/me/drive/root:/{name}.xlsx:/content".format(name=name)
        response = requests.put(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response['id']

    def add_new_worksheet(self, workbook_id, sheet_name):
        """ Add a new worksheets in a workbook.

        Parameters
        ----------
        workboook_id : str
        sheet_name : str
            
        Returns
        -------
        str
            the name of the new worksheet
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
        }
        data = {
            "name": sheet_name
        }
        url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets/".format(
            workbook_id=workbook_id)
        response = requests.post(url, json=data, headers=headers)
        json_response = json.loads(response.text)
        return json_response['name']

    def get_cell(self, workbook_id, sheet_name, range_cell):
        """ Get the value of a cell or range from a workbook.

        Parameters
        ----------
        workboook_id : str
        sheet_name : str
        range_cell: str
            
        Returns
        -------
        list
            a list of lists (representing rows and colums) with the values of the cell/s
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
        }
        url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')".format(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range_cell=range_cell
        )
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response['values']

    def update_range(self, workbook_id, sheet_name, range_cell, value_cell):
        """ Get the value of a cell or range from a workbook.

        Parameters
        ----------
        workboook_id : str
        sheet_name : str
        range_cell: str
        value_cell: str, int or float
            
        Returns
        -------
        list
            a list of lists (representing rows and colums) with the new values of the cell/s
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
        }
        cells = self.get_cell(workbook_id, sheet_name, range_cell)
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
        json_response = json.loads(response.text)
        return json_response['values']