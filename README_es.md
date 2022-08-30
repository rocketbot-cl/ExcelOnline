



# ExcelOnline
  
Trabajando con Excel Online  

## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  

## Como usar este modulo

Antes de usar este modulo, es necesario registrar tu aplicación en el portal de Azure. 

Registración: 

1. Inicie sesión en el portal de Azure. Luego diríjase al siguiente link: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade (Registración de Aplicaciones). 
2. Click en "Nueva Registración". Elija un nombre. 
3. En "Tipo de Cuentas" elija "Cuentas de cualquier directorio de la organización y cuentas personales de Microsoft (por ejemplo, Skype, Xbox)". 
4. En "URI de Redirección" selecciones "Web" como plataforma y ponga como URI: __https://localhost/__. Finalmente clieckear en "Registrar". 
5. Una vez registrada, en la sección "General" encontrara el "ID de Aplicación (Cliente)", escriba/guárdelo, lo necesitara mas adelante. 
6. Diríjase a "Certificados y Secretos", genere un nuevo "Secreto de Cliente", escriba una descripción y fije la expiración en 24 meses (preferiblemente). Click en "Adherir" y escriba/guarde el "Valor" (NO el "ID de Secreto"), lo necesitara luego junto con el ID de la App. 
7. Finalmente, vaya a "Permisos de API", clickee en "Adherir permiso", luego en "Microsoft Graph" y seleccione "Permisos delegados". En la barra de búsqueda tipee, "Files.ReadWriteAll", marque con un tilde el casillero y clickee en "Adherir Permisos". 

## Overview


1. Establecer credenciales  
Establece las credenciales para tener disponible la API

2. Obtener código de acceso  
Obtener código de acceso para generar las credenciales de la API

3. Establecer credenciales  
Establece las credenciales para tener disponible la API

4. Obtener archivos XLSX  
Devuelve una lista con todos los archivos XLSX en la locación por defecto

5. Obtener hojas de trabajo  
Obtiene hojas de trabajo de un archivo de excel

6. Crear libro  
Crear un nuevo libro en la ubicación por defecto

7. Añadir nueva hoja  
Añadir una nueva hoja al libro

8. Obtener celda  
Obtener valor de celda o rango

9. Escribir/cambiar celda  
Escribir/cambiar el valor de una celda o rango  

----
### OS

- windows
- mac
- linux

### Dependencies

### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)