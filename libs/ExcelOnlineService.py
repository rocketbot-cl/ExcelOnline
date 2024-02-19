from http import client
from re import A
import requests
import os
import msal

import json
import webbrowser
class ExcelOnlineService:
    
    def __init__(self, *, client_id, client_secret, tenant, redirect_uri= None, path_user=None, path_credentials):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant = tenant
        self.authority_url = f'https://login.microsoftonline.com/{self.tenant}/' # https://learn.microsoft.com/en-us/azure/active-directory/develop/msal-client-application-configuration 
                                                                         # I leave this here just in case is needed a Single Tenant configuration at some moment
        self.scopes = ['User.Read', 'Files.ReadWrite.All']
        self.access_token = None
        self.refresh_token = None
        self.path_user = path_user
        self.path_credentials = path_credentials
        self.base_url = 'https://graph.microsoft.com/v1.0/'
        
        # Deprecated parameters
        self.redirect_uri = redirect_uri

# --------------------------------------------------------------------------------------------------------------------------------------------------------
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
            print('---------')
            
            json_response = json.loads(response.text)
            print(json_response)
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
        if grant_type == 'refresh_token':
            params = {
                'grant_type': grant_type,
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'redirect_uri': self.redirect_uri,
            }
        else:    
            params = {
                'grant_type': grant_type,
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': "files.readwrite.all offline_access",
                'redirect_uri': self.redirect_uri,
            }
        params.update(auth_code)
        url = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token'.format(tenant=self.tenant)
        return url, params
# --------------------------------------------------------------------------------------------------------------------------------------------------------
# The funcitions over the line are deprecated. Kept for compatibility issues.
   
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
                user_info = {'client_id': self.client_id, 'client_secret': self.client_secret, 'tenant': self.tenant}
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

    def get_drives(self):
        """ Get the '.xlsx' files in the directory.
            
        Returns
        -------
        list
            a list of dictionaries with name and ID of the xlsx files
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }

        url = self.base_url + f"drives?$select=driveType,id,owner,name,description"
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            url = self.base_url + f"drives?$select=driveType,id,owner"
            response = requests.get(url, headers=headers)
        if response.status_code == 200:
            json_response = json.loads(response.text)['value']
        else:
            json_response = json.loads(response.text)
        return json_response

    def get_xlsx_files(self, folder_id = 'root', drive_id = None):
        """ Get the '.xlsx' files in the directory.
            
        Returns
        -------
        list
            a list of dictionaries with name and ID of the xlsx files
        """
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        if drive_id:
            if folder_id != 'root':
                url = self.base_url + f"/drives/{drive_id}/items/{folder_id}/search(q='.xlsx')?select=name,id,webUrl"
            else:
                url = self.base_url + f"/drives/{drive_id}/search(q='.xlsx')?select=name,id,webUrl"
        else:
            url = self.base_url + f"me/drive/items/{folder_id}/search(q='.xlsx')" #?select=name,id,webUrl
            url_shared = self.base_url + "me/drive/sharedWithMe/"
            
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        clean_data = []
        for xlsx in json_response['value']:
            xlsx_ = {}
            xlsx_['name'] = xlsx['name']
            xlsx_['id'] = xlsx['id']
            xlsx_['driveId'] = ""
            if 'parentReference' in xlsx.keys():
                xlsx_['driveId'] = xlsx['parentReference']['driveId']
            if 'remoteItem' in xlsx.keys():
                xlsx_['driveId'] = xlsx['remoteItem']['parentReference']['driveId']
            
            clean_data.append(xlsx_)
            
            # for k in ['@odata.type','webUrl']:
            #     xlsx.pop(k)
            # clean_data.append(xlsx)
        
        if url_shared:
            response_shared = requests.get(url_shared, headers=headers)
            json_response_shared = json.loads(response_shared.text)
            for file in json_response_shared['value']:
                if '.xlsx' in file['name']:
                    file_ = {}
                    file_['name'] = file['name']
                    file_['id'] = file['id']
                    try:
                        file_['driveId'] = file['parentReference']['driveId']
                    except:
                        file_['driveId'] = file['remoteItem']['parentReference']['driveId']
                    
                    clean_data.append(file_)
        
        return clean_data
        
    def get_worksheets(self, workbook_id, drive_id = None):
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
        if drive_id:
            url = self.base_url + f"/drives/{drive_id}/items/{workbook_id}/workbook/worksheets"
        else:
            url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets"
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        sheets = []
        for sheet in json_response['value']:
            sheets.append(sheet['name'])
        return sheets

    # This function was replaced by upload_item, because to make the workbook available, it was neccesary to get into OneDrive and open it.    
    # # def create_workbook(self, name):
    # #     """ Creates a new workbook.

    # #     Parameters
    # #     ----------
    # #     name : str
            
    # #     Returns
    # #     -------
    # #     str
    # #         the ID of the new workbook
    # #     """
    # #     headers = {
    # #         'Authorization': 'Bearer ' + self.access_token
    # #     }
    # #     url = self.base_url + f"/me/drive/root:/{name}.xlsx:/content".format(name=name)
    # #     response = requests.put(url, headers=headers)
    # #     json_response = json.loads(response.text)
    # #     return json_response['id']
    
    def upload_item(self, file_path, filename, drive_id = None, folder_id = None):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        
        # Open file in binary format for reading to upload it
        with open(file_path, "rb") as file:
            fileHandle = file.read()
            
        if drive_id and folder_id:
            url = self.base_url + f"/drives/{drive_id}/items/{folder_id}:/{filename}:/content"
        else:
            url = self.base_url + f"/me/drive/items/root/{filename}:/content"
            
        response = requests.put(url, data=fileHandle, headers=headers)
        json_response = json.loads(response.text)

        return json_response['id']
       
    def create_session(self, id, drive_id = None):
        headers = {
            'Authorization': 'Bearer ' + self.access_token
        }
        session_params = {
            "persistChanges": True
        }
        if drive_id:
            url = self.base_url + f"/drives/{drive_id}/items/{id}/workbook/createSession"
        else:
            url = self.base_url + f"/me/drive/items/{id}/workbook/createSession"
        response = requests.post(url, json=session_params, headers=headers)
        json_response = json.loads(response.text)
        print(json_response)
        return json_response['id']
    
    def close_session(self, id, session_id, drive_id = None):
        headers = {
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        if drive_id:
            url = self.base_url + f"/drives/{drive_id}/items/{id}/workbook/closeSession"
        else:
            url = self.base_url + f"/me/drive/items/{id}/workbook/closeSession"
        
        response = requests.post(url, headers=headers)

        return response

    def add_new_worksheet(self, workbook_id, sheet_name, session_id, drive_id = None):
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
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        data = {
            "name": sheet_name
        }
        if drive_id:
            url = self.base_url + f"/drives/{drive_id}/items/{workbook_id}/workbook/worksheets/"
        else:
            url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets/"
            
        response = requests.post(url, json=data, headers=headers)
        json_response = json.loads(response.text)
        return json_response['name']

    def get_cell(self, workbook_id, sheet_name, range_cell, session_id, drive_id = None):
        """ Get the value of a cell or range from a workbook.

        Parameters
        ----------
        workboook_id : str
        sheet_name : str
        range_cell: str
            
        Returns
        -------
        list
            a list of lists (representing rows and columns) with the values of the cell/s
        """
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        if drive_id:
            url = self.base_url + f"/drives/{drive_id}/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')"
        else:
            url = self.base_url + f"/me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')"
            
        response = requests.get(url, headers=headers)
        json_response = json.loads(response.text)
        return json_response['values']

    def update_range(self, workbook_id, sheet_name, range_cell, value_cell, session_id, drive_id = None):
        """ Get the value of a cell or range from a workbook.

        Parameters
        ----------
        workboook_id : str
        sheet_name : str
        range_cell: str
        value_cell: str, int, float or list
            
        Returns
        -------
        list
            a list of lists (representing rows and columns) with the new values of the cell/s
        """
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        
        # It takes the input and parses it (from str) into to its type (int, float, list, etc)
        try:
            new_values = eval(value_cell)
        except:
            new_values = value_cell
            
        data = {
            "values" : new_values
        }

        if drive_id:
            url = self.base_url + f"/drives/{drive_id}/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')"
        else:
            url = self.base_url + f"me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')"
        
        response = requests.patch(url, json=data, headers=headers)
        json_response = json.loads(response.text)
        
        # It checks that the modification was done.
        try:
            if new_values == json_response['values'] or new_values in json_response['values'][0]:
                return json_response['values']
            else:
                raise ValueError("Check values matrix structure, may not fit given range.")
        except:
            raise ValueError("Check values matrix structure, may not fit given range.")
        
    def write_formula(self, workbook_id, sheet_name, range_cell, formula, session_id, drive_id = None):
        """ Get the value of a cell or range from a workbook.

        Parameters
        ----------
        workboook_id : str
        sheet_name : str
        range_cell: str
        formula: str
            
        Returns
        -------
        list
            a list of lists (representing rows and columns) with the new values of the cell/s
        """
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + self.access_token,
            'workbook-session-id': session_id
        }
        
        data = {
            "formulas" : formula
        }

        if drive_id:
            url = self.base_url + f"/drives/{drive_id}/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')"
        else:
            url = self.base_url + f"me/drive/items/{workbook_id}/workbook/worksheets/{sheet_name}/range(address='{range_cell}')"
        
        response = requests.patch(url, json=data, headers=headers)
        json_response = json.loads(response.text)
        
        return json_response
