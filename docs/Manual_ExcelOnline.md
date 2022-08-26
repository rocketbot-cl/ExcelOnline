



# ExcelOnline
  
Working with Excel Online  
  
![banner](/docs/imgs/Banner_C:\Users\jmsir\Desktop\RB\Rocketbot\modules\ExcelOnline.png)
## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  


## Como usar este modulo

Antes de usar este modulo, es necesario registrar tu aplicación en el portal de Azure. 


Registración: 

1. Inicie sesión en el portal de Azure. Luego diríjase al siguiente link: 
https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade (Registración de Aplicaciones). 
2. 
Click en "Nueva Registración". Elija un nombre. 
3. En "Tipo de Cuentas" elija "Cuentas de cualquier directorio de la 
organización y cuentas personales de Microsoft (por ejemplo, Skype, Xbox)". 
4. En "URI de Redirección" selecciones 
"Web" como plataforma y ponga como URI: __https://localhost/__. Finalmente clieckear en "Registrar". 
5. Una vez 
registrada, en la sección "General" encontrara el "ID de Aplicación (Cliente)", escriba/guárdelo, lo necesitara mas 
adelante. 
6. Diríjase a "Certificados y Secretos", genere un nuevo "Secreto de Cliente", escriba una descripción y fije
 la expiración en 24 meses (preferiblemente). Click en "Adherir" y escriba/guarde el "Valor" (NO el "ID de Secreto"), lo
 necesitara luego junto con el ID de la App. 
7. Finalmente, vaya a "Permisos de API", clickee en "Adherir permiso", 
luego en "Microsoft Graph" y seleccione "Permisos delegados". En la barra de búsqueda tipee, "Files.ReadWriteAll", 
marque con un tilde el casillero y clickee en "Adherir Permisos". 


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
|Cell or range|Cell or range|A1:B1|
|Variable to assign|Variable to assign. Returns cell or range value/es|valor_celda|

### Write/change cell
  
Write/change a cell or range value
|Parameters|Description|example|
| --- | --- | --- |
|Workbook ID|Workbook ID|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Worksheet name|Worksheet name|Sheet1|
|Cell or range|Cell or range where it will be written|A1:B1|
|Value cell or range|Value of the cell or range|value_cell|
|Variable to assign|Variable to assign. Returns True if the change was successful, otherwise it will be False|nuevo_valor_celda|
