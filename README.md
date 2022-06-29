



# ExcelOnline
  
Working with Excel Online  

## Howto install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  


## Como usar este modulo

Before using this module, you must register your app into the Azure Portal.

Registrations.

1. Sign in to Azure Portal (App Registrations). Then go to the next link: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
2. Click on "New registration". Choose a name.
3. In “Supported account types” choose "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)".
4. In "Redirect URI" select "Web" as a plataform an set the URI to: https://localhost/. Finally click on "Register"
5. Once registered, on the "Overview" section you will find the "Application (client) ID", write it down/save it, you will need it later.
6. Go to "Certificates & secrets", generate a "New client secret", write a description and set the expration time to 24 months. Click on "Add" and write down/save the "Value" (NOT the "Secret ID"), you will need it later together with the App ID.
7. Finally, go to "API persmissions", click on "Add a permission", then on "Microsoft Graph" and select "Permisos delegados". In the serch bar type Files.ReadWriteAll, mark the checkbox and click on "Add permissions".
8. Get the Access Code with "Get Access Code" funtion. Using the "Application (client) ID" and the "Secret Value", after clicking on accept the web browser will open asking to sign in and then to accept the conection with your application. Once clicked on "Continue" it will redirect you to a URL like this: https://localhost/?code=M.R3_BAY.c0173ccb-c865-d99f-008c-2c7c9478d63f. Copy whats after "code=", and use it in Set Credentials.
9. Use the code to generate the credentials for the first time with __Set credentials__. Once generated, there will not be needed to do the previous steps again until the Client Secret (obtained in step 6) expires.


## Overview


1. Get access code
Get access code

2. Set credentials
Set credentials to make available the API

3. Get XLSX files
Return a list with all the XLSX files in the default location

4. Get worksheets
Get worksheets from an excel file

5. Create workbook
Create a new workbook in the default location. 
Note: It does not replace an existing workbook.

6. Add new worksheet
Add a new worksheet to the workbook
Note: It does not replace an existing sheet.

7. Get cell
Get cell value

8. Write cell
Write/change a cell value

### Changes
Mon Dec 27 11:09:45 2021  (refs/stash) On master: !!GitHub_Desktop<master>

----
### OS

- windows
- mac
- linux

### Dependencies

### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)