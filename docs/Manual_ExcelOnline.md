



# ExcelOnline
  
Create, read, update and delete data from Excel files stored into OneDrive  

*Read this in other languages: [English](Manual_ExcelOnline.md), [Português](Manual_ExcelOnline.pr.md), [Español](Manual_ExcelOnline.es.md)*
  
![banner](imgs/Banner_ExcelOnline.png)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  




## How to use this module

Before using this module, you need to register your app in the Azure App Registrations portal.

1. Sign in to the Azure portal and search for the Azure Active Directory service.
2. On the left side menu, get into "App Registrations".
3. Select "New record".
4. Under “Compatible account types” supported choose:
    - "Accounts in any organizational directory (any Azure AD directory: multi-tenant) and personal Microsoft accounts (such as Skype or Xbox)" for this case use Tenant ID = **common**.
    - "Only accounts from this organizational directory (only this account: single tenant) for this case use application-specific **Tenant ID**.
    - "Personal Microsoft accounts only" for this case use use Tenant ID = **consumers**.
5. In "Redirect URI" select "Web" as a plataform and set the URI to: https://localhost/. Finally click on "Register"
6. Copy the application (client) ID. You will need this value.
7. Under "Certificates and secrets", generate a new 
client secret. Set the expiration (preferably 24 months). Copy the **VALUE** of the created client secret (**__NOT the Secret ID__**). It will hide after a few minutes.
8. Finally, go to "API persmissions", click "Add a permission", select "Microsoft Graph", then "Delegated permissions". In the search bar type "Files.ReadWriteAll", mark the checkbox and click on "Add permissions".


## Description of the commands

### Get access code
  
Get access code to create the credentials for the API
|Parameters|Description|example|
| --- | --- | --- |
|client_id|Client ID from the API|cliente_id|
|client_secret_value|Client Secret Value from the API|valor_secreto_cliente|
|tenant|API Tenant ID|common|

### Set credentials
  
Set credentials to make available the API
|Parameters|Description|example|
| --- | --- | --- |
|code|Authorization code|code|
|Variable to assign|Variable to assign. If the connection is successful, it will return True, otherwise it will return False|variable|

### Get XLSX files
  
Return a list with all the XLSX files in 
|Parameters|Description|example|
| --- | --- | --- |
|Shared Drive ID (Optional)||097JB2CA2559D776|
|Folder ID (Optional)|||
|Variable to assign|Variable to assign. Returns the list of files|lista_archivos|

### Get worksheets
  
Get worksheets from an excel file
|Parameters|Description|example|
| --- | --- | --- |
|Shared Drive ID (Optional)||097JB2CA2559D776|
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Variable to assign|Variable to assign. Returns the list of worksheets of the workbook|lista_hojas|

### Create workbook
  
Create a new workbook in the default location
|Parameters|Description|example|
| --- | --- | --- |
|Shared Drive ID (Optional)||097JB2CA2559D776|
|Workbook name|Workbook name|Libro Nuevo|
|Variable to assign|Variable to assign. Returns the ID of the new workbook|id_nuevoLibro|

### Add new worksheet
  
Add a new worksheet to the workbook
|Parameters|Description|example|
| --- | --- | --- |
|Shared Drive ID (Optional)||097JB2CA2559D776|
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Variable to assign|Variable to assign. Returns the name of the new sheet|nombre_nuevaHoja|

### Get cell
  
Get a cell or range values
|Parameters|Description|example|
| --- | --- | --- |
|Shared Drive ID (Optional)||097JB2CA2559D776|
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Cell or range|Cell or range A1B2|Cell or Range|
|Variable to assign|Variable to assign. Returns cell or range value/es|valor_celda|

### Write/change cell
  
Write/change a cell or range value
|Parameters|Description|example|
| --- | --- | --- |
|Shared Drive ID (Optional)||097JB2CA2559D776|
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Cell or range|Cell or range where it will be written A1B2|Cell or Range|
|Value cell or range|Value of the cell or range [[1,2],[1,2]]|Value or list of values|
|Variable to assign|Variable to assign. Returns True if the change was successful, otherwise it will be False|nuevo_valor_celda|

### Write formula
  
Write a formula into a cell or range
|Parameters|Description|example|
| --- | --- | --- |
|Shared Drive ID (Optional)||097JB2CA2559D776|
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Cell or range|Cell or range where it will be written A1B2|Cell or Range|
|Formula|Excel format Formula|=sum(2,2)|
|Variable to assign|Variable to assign. Returns True if the change was successful, otherwise it will be False|nuevo_valor_celda|

### (DEPRECATED) Set credentials
  
Set credentials to make available the API
|Parameters|Description|example|
| --- | --- | --- |
|client_id|Client ID from the API|Your client_id|
|client_secret|Client Secret from the API|Your client_secret|
|redirect_uri|Redirection of the API|http://localhost:5000|
|code|Authorization code|code|
|tenant|Tenant of the API|tenant|
|connection|Variable where we will store our result. If the connection is successful, it will return True, otherwise it will return False|connection|
