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
from time import time, sleep
import openpyxl
import traceback
base_path = tmp_global_obj["basepath"]
cur_path = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + "libs" + os.sep
if cur_path not in sys.path:
    sys.path.append(cur_path)

class NoCode(Exception):
    """Raise when no code have been generated for a new connection"""
    pass

# Get the module that was called
module = GetParams("module")

from ExcelOnlineService import ExcelOnlineService

global excel_online_service, _client

# Set the paths for the output files
user_filename = 'user.json'
credentials_filename = 'credentials.json'
path_user = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + user_filename
path_credentials = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + credentials_filename

"""
--------------------------------------------------------------------------------------------------------------------------------------------------------
"""
if module == "setCredentials":
    client_secret = GetParams("client_secret")
    client_id = GetParams("client_id")
    redirect_uri = GetParams("redirect_uri")
    code = GetParams("code")
    tenant = GetParams("tenant")
    res = GetParams("res")
    credentials_filename = 'credentials.json'
    excel_online_service = ExcelOnlineService(client_id=client_id, client_secret=client_secret, tenant=tenant, redirect_uri=redirect_uri,
                        path_credentials=path_credentials)
    try:
        try:
            with open(path_credentials) as json_file:
                data = json.load(json_file)
            auth_code = {'refresh_token': data['refresh_token']}
            print(auth_code)
            grant_type = 'refresh_token'
            response = excel_online_service.get_token(auth_code, grant_type)
        except IOError:
            grant_type = 'authorization_code'
            auth_code = {'code': code}
            print(auth_code)
            response = excel_online_service.get_token(auth_code, grant_type)
            print(response)
        is_connected = excel_online_service.create_tokens_file(response)
        SetVar(res, is_connected)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e
"""
--------------------------------------------------------------------------------------------------------------------------------------------------------
The funcitions used between the line are deprecated. Kept for compatibility issues.
"""

if module == "getCode":
    # Creates the instance of a client and obtains the authorization code
    client_id = GetParams("client_id")
    client_secret = GetParams("client_secret")
    tenant = GetParams("tenant")

    if not tenant:
        tenant = "common"
    
    excel_online_service = ExcelOnlineService(client_id=client_id, client_secret=client_secret, tenant=tenant, path_user=path_user,
                    path_credentials=path_credentials)

    _client = excel_online_service.get_code()
    
if module == "setCredentials_2":
    # Seeks for the refresh_token in the credentials and client data, if does not exist, it will take the client instance and the generates code to create the credentials.
    code = GetParams("code")
    res = GetParams("res")
    try:
        try:
            with open(path_credentials) as cred_file:
                credentials = json.load(cred_file)
            with open(path_user) as user_file:
                client_data = json.load(user_file)
            
            if client_data.get('tenant'):
                tenant = client_data['tenant']
            else:
                tenant = "common"
                            
            excel_online_service = ExcelOnlineService(client_id=client_data['client_id'], client_secret=client_data['client_secret'], tenant=tenant, path_user=path_user,
                    path_credentials=path_credentials)
                
            refresh_token = credentials['refresh_token']
            response = excel_online_service.get_old_token(refresh_token)
            SetVar(res, True)
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
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e

if module == "get_xlsx_files":
    drive_id = GetParams("drive_id")
    res = GetParams("res")
    try:
        if drive_id:
            files = excel_online_service.get_xlsx_files(drive_id)
        else:
            files = excel_online_service.get_xlsx_files()
        SetVar(res, files)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e

if module == "get_worksheets":
    drive_id = GetParams("drive_id")
    res = GetParams("res")
    workbook_id = GetParams("workbook_id")
    try:
        if drive_id:
            session_id = excel_online_service.create_session(workbook_id, drive_id)
            list_worksheets = excel_online_service.get_worksheets(workbook_id, drive_id)      
            excel_online_service.close_session(workbook_id, session_id, drive_id)
        else:
            session_id = excel_online_service.create_session(workbook_id)
            list_worksheets = excel_online_service.get_worksheets(workbook_id)
            excel_online_service.close_session(workbook_id, session_id)
        SetVar(res, list_worksheets)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e

if module == "create_workbook":
    drive_id = GetParams("drive_id")
    res = GetParams("res")
    workbook_name = GetParams("workbook_name")

    """It is not possible to create a Workbook directly in the cloud an use it right after, necessarily the user must manually open it or wait several ours so the
    bot can make use of it. So, to avoid that, the workbook is created locally, the upload and finally erased from the computer."""
    # It creates a temporary .xlsx file in 'tmp' folder
    path_temp = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + "tmp"
    if not os.path.exists(path_temp):
        os.makedirs(path_temp)
    workbook_name_ = workbook_name + ".xlsx"
    temp_file = os.path.join(path_temp, workbook_name_)
    
    wb = openpyxl.Workbook()
    wb.save(temp_file)
    
    try:
        if drive_id:
            data_new_workbook = excel_online_service.upload_item(temp_file, workbook_name_, drive_id)
        else:
            data_new_workbook = excel_online_service.upload_item(temp_file, workbook_name_)
        SetVar(res, True)
        sleep(20)
        # Once the file is uploaded, it is erased
        os.remove(temp_file)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e

if module == "add_new_worksheet":
    drive_id = GetParams("drive_id")
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    res = GetParams("res")
    try:
        if drive_id:
            session_id = excel_online_service.create_session(workbook_id, drive_id)
            new_sheet = excel_online_service.add_new_worksheet(workbook_id, worksheet_name, session_id, drive_id)
            excel_online_service.close_session(workbook_id, session_id, drive_id)
        else:
            session_id = excel_online_service.create_session(workbook_id)
            new_sheet = excel_online_service.add_new_worksheet(workbook_id, worksheet_name, session_id)
            excel_online_service.close_session(workbook_id, session_id)        
        SetVar(res, True)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e

if module == "get_cell":
    drive_id = GetParams("drive_id")
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    range_cell = GetParams("range_cell")
    res = GetParams("res")
    try:
        if drive_id:
            session_id = excel_online_service.create_session(workbook_id, drive_id)
            value_cell = excel_online_service.get_cell(workbook_id, worksheet_name, range_cell, session_id, drive_id)
            excel_online_service.close_session(workbook_id, session_id, drive_id)
        else:
            session_id = excel_online_service.create_session(workbook_id)
            value_cell = excel_online_service.get_cell(workbook_id, worksheet_name, range_cell, session_id)
            excel_online_service.close_session(workbook_id, session_id)
        SetVar(res, value_cell)
        a = excel_online_service.close_session(session_id)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e

if module == "update_range":
    drive_id = GetParams("drive_id")
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    range_cell = GetParams("range_cell")
    value_cell = GetParams("value_cell")
    res = GetParams("res")
    try:
        if drive_id:
            session_id = excel_online_service.create_session(workbook_id, drive_id)
            new_value_cell = excel_online_service.update_range(workbook_id, worksheet_name, range_cell, value_cell, session_id, drive_id)
            excel_online_service.close_session(workbook_id, session_id, drive_id)
        else:
            session_id = excel_online_service.create_session(workbook_id)
            new_value_cell = excel_online_service.update_range(workbook_id, worksheet_name, range_cell, value_cell, session_id)
            excel_online_service.close_session(workbook_id, session_id)
        SetVar(res, True)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e
    
if module == "write_formula":
    drive_id = GetParams("drive_id")
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    range_cell = GetParams("range_cell")
    formula = GetParams("formula")
    res = GetParams("res")
    try:
        if drive_id:
            session_id = excel_online_service.create_session(workbook_id, drive_id)
            new_value_cell = excel_online_service.write_formula(workbook_id, worksheet_name, range_cell, formula, session_id, drive_id)
            excel_online_service.close_session(workbook_id, session_id, drive_id)
        else:
            session_id = excel_online_service.create_session(workbook_id)
            new_value_cell = excel_online_service.write_formula(workbook_id, worksheet_name, range_cell, formula, session_id)
            excel_online_service.close_session(workbook_id, session_id)
        SetVar(res, True)
    except Exception as e:
        SetVar(res, False)
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        PrintException()
        traceback.print_exc()
        raise e