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


base_path = tmp_global_obj["basepath"]
cur_path = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + "libs" + os.sep
if cur_path not in sys.path:
    sys.path.append(cur_path)

import traceback
from ExcelOnlineService import ExcelOnlineService

"""
    Obtengo el modulo que fue invocado
"""
global excel_online_service
module = GetParams("module")

if module == "setCredentials":
    client_secret = GetParams("client_secret")
    client_id = GetParams("client_id")
    redirect_uri = GetParams("redirect_uri")
    code = GetParams("code")
    tenant = GetParams("tenant")
    res = GetParams("res")
    credentials_filename = 'credentials.json'
    path_credentials = base_path + "modules" + os.sep + "ExcelOnline" + os.sep + credentials_filename
    excel_online_service = ExcelOnlineService(client_id=client_id, client_secret=client_secret, tenant=tenant, redirect_uri=redirect_uri,
                        path_credentials=path_credentials)
    try:
        try:
            with open(path_credentials) as json_file:
                data = json.load(json_file)
            auth_code = {'refresh_token': data['refresh_token']}
            grant_type = 'refresh_token'
            response = excel_online_service.get_token(auth_code, grant_type)
        except IOError:
            grant_type = 'authorization_code'
            auth_code = {'code': code}
            response = excel_online_service.get_token(auth_code, grant_type)
        is_connected = excel_online_service.create_tokens_file(response)
        SetVar(res, is_connected)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        #print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e

if module == "get_xlsx_files":
    res = GetParams("res")
    try:
        files = excel_online_service.get_xlsx_files()
        SetVar(res, files['value'])
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e

if module == "get_worksheets":
    res = GetParams("res")
    workbook_id = GetParams("workbook_id")
    try:
        list_worksheets = excel_online_service.get_worksheets(workbook_id)
        print(list_worksheets)
        SetVar(res, list_worksheets['value'])
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        print("Traceback: ", traceback.format_exc())
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
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e

if module == "add_new_worksheet":
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    try:
        session_id = excel_online_service.create_session(workbook_id)['id']
        excel_online_service.add_new_worksheet(session_id, workbook_id, worksheet_name)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e

if module == "get_cell":
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    range_cell = GetParams("range_cell")
    res = GetParams("res")
    try:
        session_id = excel_online_service.create_session(workbook_id)['id']
        value_cell = excel_online_service.get_cell(session_id, workbook_id, worksheet_name, range_cell)
        SetVar(res, value_cell)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e

if module == "update_range":
    workbook_id = GetParams("workbook_id")
    worksheet_name = GetParams("worksheet_name")
    range_cell = GetParams("range_cell")
    value_cell = GetParams("value_cell")
    try:
        session_id = excel_online_service.create_session(workbook_id)['id']
        excel_online_service.update_range(session_id, workbook_id, worksheet_name, range_cell, value_cell)
    except Exception as e:
        print("\x1B[" + "31;40mAn error occurred\x1B[" + "0m")
        print("Traceback: ", traceback.format_exc())
        PrintException()
        raise e