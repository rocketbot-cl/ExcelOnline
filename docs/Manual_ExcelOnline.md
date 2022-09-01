# ExcelOnline
  
Working with Excel Online  
  
![banner](imgs/Banner_ExcelOnline.png)
## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  

## Como usar este modulo

## How to use this module

Before using this module, you must register your app into the Azure Portal.


Registration:

1. Sign in to Azure Portal. Then go to the next link: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade (App Registrations).
2. Click on "New registration". Choose a name.
3. In “Supported account types” choose "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)".
4. In "Redirect URI" select "Web" as a plataform and set the URI to: https://localhost/. Finally click on "Register"
5. Once registered, on the "Overview" section you will find the "Application (client) ID", write it down/save it, you will need it later.
6. Go to "Certificates & secrets", generate a "New client secret", write a description and set the expration time to 24 months (preferably). Click on "Add" and write down/save the "Value" (NOT the "Secret ID"), you will need it later together with the App ID.
7. Finally, go to "API persmissions", click on "Add a permission", then on "Microsoft Graph" and select "Delegated permissions". In the search bar type "Files.ReadWrite.All", mark the checkbox and click on "Add permissions".

## Description of the commands

### Set credentials
  
Set credentials to make available the API
|Parameters|Description|example|
| --- | --- | --- |
|client_id|Client ID from the API|Your client_id|
|client_secret|Client Secret from the API|Your client_secret|
|redirect_uri|Redirection of the API|http://localhost:5000|
|code|Authorization code|code|
|tenant|Tenant of the API|tenant|
|connection|Variable where we will store our result. If the connection is successful, it will return True, otherwise it will return False|connection|

### Get access code
  
Get access code to create the credentials for the API
|Parameters|Description|example|
| --- | --- | --- |
|client_id|Client ID from the API|cliente_id|
|client_secret_value|Client Secret Value from the API|valor_secreto_cliente|

### Set credentials
  
Set credentials to make available the API
|Parameters|Description|example|
| --- | --- | --- |
|code|Authorization code|code|
|Variable to assign|Variable to assign. If the connection is successful, it will return True, otherwise it will return False|variable|

### Get XLSX files
  
Return a list with all the XLSX files in the default location
|Parameters|Description|example|
| --- | --- | --- |
|Variable to assign|Variable to assign. Returns the list of files|lista_archivos|

### Get worksheets
  
Get worksheets from an excel file
|Parameters|Description|example|
| --- | --- | --- |
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Variable to assign|Variable to assign. Returns the list of worksheets of the workbook|lista_hojas|

### Create workbook
  
Create a new workbook in the default location
|Parameters|Description|example|
| --- | --- | --- |
|Workbook name|Workbook name|Libro Nuevo|
|Variable to assign|Variable to assign. Returns the ID of the new workbook|id_nuevoLibro|

### Add new worksheet
  
Add a new worksheet to the workbook
|Parameters|Description|example|
| --- | --- | --- |
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Variable to assign|Variable to assign. Returns the name of the new sheet|nombre_nuevaHoja|

### Get cell
  
Get a cell or range values
|Parameters|Description|example|
| --- | --- | --- |
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Cell or range|Cell or range: A1:B2|Cell or Range|
|Variable to assign|Variable to assign. Returns cell or range value/es|valor_celda|

### Write/change cell
  
Write/change a cell or range value
|Parameters|Description|example|
| --- | --- | --- |
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Cell or range|Cell or range where it will be written: A1:B2|Cell or Range|
|Value cell or range|Value of the cell or range: [[1,2],[1,2]]|Value or list of values|
|Variable to assign|Variable to assign. Returns True if the change was successful, otherwise it will be False|nuevo_valor_celda|
