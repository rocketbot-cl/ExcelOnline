



# ExcelOnline
  
Working with Excel Online  

## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  


## Como usar este modulo

Para permitir la autenticación, primero debe registrar su aplicación en Azure App 
Registrations.

1- Inicie sesión en Azure Portal (App Registrations)
2- Cree una aplicación. Establezca un nombre.
3- En
 “Tipos de cuenta” soportados elige "Cuentas de cualquier directorio de la organización y cuentas personales de 
Microsoft (por ejemplo, Skype, Xbox, Outlook.com)".
4- Establezca la uri de redirección (Web) : https://localhost y haga
 click en registrarse.
5- Anote el ID de la aplicación (cliente). Necesitará este valor.
6- En "Certificados y 
secretos", genere un nuevo secreto de cliente. Establezca que la caducidad sea preferiblemente 24 meses. Anote el VALOR 
del secreto de cliente creado ahora. Se ocultará más adelante. Debe copiar VALOR, no Id de secreto.
7- En Permisos de 
API, añadir los siguientes permisos, en permisos delegados y de aplicación: Files.ReadWriteAll



## Overview


1. Set credentials  
Set credentials to make available the API

2. Get XLSX files  
Return a list with all the XLSX files in the folder

3. Get worksheets  
Get worksheets from an excel file

4. Create workbook  
Create a new workbook in default location

5. Add new worksheet  
Add a new worksheet to the workbook

6. Get cell  
Get cell value

7. Write cell  
Write a cell in a sheet  




----
### OS

- windows
- mac
- linux

### Dependencies

### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)