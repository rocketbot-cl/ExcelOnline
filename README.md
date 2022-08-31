



# ExcelOnline
  
Working with Excel Online  

## Howto install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  

## How to use this module

Before using this module, you must register your app into the Azure Portal.

Registration:

1. Sign in to Azure Portal. Then go to the next link: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade (App Registrations).
2. Click on "New registration". Choose a name.
3. In “Supported account types” choose "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)".
4. In "Redirect URI" select "Web" as a plataform and set the URI to: https://localhost/. Finally click on "Register"
5. Once registered, on the "Overview" section you will find the "Application (client) ID", write it down/save it, you will need it later.
6. Go to "Certificates & secrets", generate a "New client secret", write a description and set the expration time to 24 months (preferably). Click on "Add" and write down/save the "Value" (NOT the "Secret ID"), you will need it later together with the App ID.
7. Finally, go to "API persmissions", click on "Add a permission", then on "Microsoft Graph" and select "Delegated permissions". In the search bar type "Files.ReadWriteAll", mark the checkbox and click on "Add permissions".

## Overview


1. Set credentials  
Set credentials to make available the API

2. Get access code  
Get access code to create the credentials for the API

3. Set credentials  
Set credentials to make available the API

4. Get XLSX files  
Return a list with all the XLSX files in the default location

5. Get worksheets  
Get worksheets from an excel file

6. Create workbook  
Create a new workbook in the default location

7. Add new worksheet  
Add a new worksheet to the workbook

8. Get cell  
Get a cell or range values

9. Write/change cell  
Write/change a cell or range value  

----
### OS

- windows
- mac
- linux

### Dependencies

### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)
