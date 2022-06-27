# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
import json
import sys
import os

base_path = tmp_global_obj["basepath"]
cur_path = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + "libs" + os.sep
if cur_path not in sys.path:
    sys.path.append(cur_path)

class NoCode(Exception):
    """Raise when no code have been generated for a new connection"""
    pass
    

from ExcelOnlineService import ExcelOnlineService

global excel_online_service, _client

# Set the paths for the output files
user_filename = 'user.json'
credentials_filename = 'credentials.json'
path_user = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + user_filename
path_credentials = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + credentials_filename


# Get the module that was called
module = GetParams("module")

if module == "getCode":
    # Creates the instance of a client and obtains the authorization code
    client_id = GetParams("client_id")
    client_secret = GetParams("client_secret")

    excel_online_service = ExcelOnlineService(client_id=client_id, client_secret=client_secret, path_user=path_user,
                    path_credentials=path_credentials)

    _client = excel_online_service.get_code()
    
if module == "setCredentials":
    # Seeks for the refresh_token in the credentials and client data, if does not exist, it will take the client instance and the generates code to create the credentials.
    code = GetParams("code")
    res = GetParams("res")
    try:
        try:
            with open(path_credentials) as cred_file:
                credentials = json.load(cred_file)
            with open(path_user) as user_file:
                client_data = json.load(user_file)
            
            excel_online_service = ExcelOnlineService(client_id=client_data['client_id'], client_secret=client_data['client_secret'], path_user=path_user,
                    path_credentials=path_credentials)
                
            refresh_token = credentials['refresh_token']
            response = excel_online_service.get_old_token(refresh_token)
        except IOError:
            if os.path.exists(path_user):
                client_instance = _client
                auth_code = code
                response = excel_online_service.get_new_token(client_instance, auth_code)
                is_connected = excel_online_service.create_tokens_file(response)
                SetVar(res, is_connected)
            else: 
                raise NoCode('Code not generated. Generate the code and then try the new connection.')
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "get_xlsx_files":
    res = GetParams("res")
    try:
        files = excel_online_service.get_xlsx_files()
        SetVar(res, files)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "get_worksheets":
    res = GetParams("res")
    workbook_id = GetParams("workbook_id")
    try:
        list_worksheets = excel_online_service.get_worksheets(workbook_id)
        SetVar(res, list_worksheets)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "create_workbook":
    res = GetParams("res")
    workbook_name = GetParams("workbook_name")
    try:
        data_new_workbook = excel_online_service.create_workbook(workbook_name)
        SetVar(res, data_new_workbook)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "add_new_worksheet":
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    res = GetParams("res")
    try:
        new_sheet = excel_online_service.add_new_worksheet(workbook_id, worksheet_name)
        SetVar(res, new_sheet)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "get_cell":
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    range_cell = GetParams("range_cell")
    res = GetParams("res")
    try:
        value_cell = excel_online_service.get_cell(workbook_id, worksheet_name, range_cell)
        SetVar(res, value_cell)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e

if module == "update_range":
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    range_cell = GetParams("range_cell")
    value_cell = GetParams("value_cell")
    res = GetParams("res")
    try:
        new_value_cell = excel_online_service.update_range(workbook_id, worksheet_name, range_cell, value_cell)
        SetVar(res, new_value_cell)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        raise e