



# ExcelOnline
  
Trabajando con Excel Online  
  
![banner](imgs/Banner_ExcelOnline.png)
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
7. Finalmente, vaya a "Permisos de API", clickee en "Adherir permiso", luego en "Microsoft Graph" y seleccione "Permisos delegados". En la barra de búsqueda tipee, "Files.ReadWrite.All", marque con un tilde el casillero y clickee en "Adherir Permisos". 


## Descripción de los comandos

### Establecer credenciales
  
Establece las credenciales para tener disponible la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|client_id|Client ID de la API|Your client_id|
|client_secret|Client Secret de la API|Your client_secret|
|redirect_uri|Redirección de la API|http://localhost:5000|
|code|Código de autorización|code|
|tenant|Tenant de la API|tenant|
|conexion|Variable donde guardaremos nuestro resultado. Si la conexion es exitosa retornara True, caso contraria sera False|connection|

### Obtener código de acceso
  
Obtener código de acceso para generar las credenciales de la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|id_cliente|ID de Cliente de la API|cliente_id|
|valor_secreto_cliente|Valor del Secreto del Cliente de la API|valor_secreto_cliente|

### Establecer credenciales
  
Establece las credenciales para tener disponible la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|code|Código de autorización|code|
|Variable a asignar|Variable a asignar. Si la conexion es exitosa retorna True, caso contraria sera False|variable|

### Obtener archivos XLSX
  
Devuelve una lista con todos los archivos XLSX en la locación por defecto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Variable a asignar|Variable a asignar. Retorna la lista de archivos|lista_archivos|

### Obtener hojas de trabajo
  
Obtiene hojas de trabajo de un archivo de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del Libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Variable a asignar|Variable a asignar. Retorna la lista de hojas del libro|lista_hojas|

### Crear libro
  
Crear un nuevo libro en la ubicación por defecto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre del libro|Nombre del libro|Libro Nuevo|
|Variable a asignar|Variable a asignar. Retorna ID del nuevo libro|id_nuevoLibro|

### Añadir nueva hoja
  
Añadir una nueva hoja al libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Variable a asignar|Variable a asignar. Retorna el nombre de la nueva hoja|nombre_nuevaHoja|

### Obtener celda
  
Obtener valor de celda o rango
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del Libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Celda o rango|Celda o rango: A1:B2|Cell or Range|
|Variable a asignar|Variable a asignar. Retorna el valor de la celda o rango|valor_celda|

### Escribir/cambiar celda
  
Escribir/cambiar el valor de una celda o rango
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Celda o rango|Celda o rango donde se escribira: A1:B2|Cell or Range|
|Valor celda o rango|Valor de la celda o rango: [[1,2],[1,2]]|Value or list of values|
|Variable a asignar|Variable a asignar. Retorna True si el cambio fue exitoso, caso contraria sera False|nuevo_valor_celda|