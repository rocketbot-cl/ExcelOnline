# ExcelOnline
  
![banner](C:\Users\jmsir\Desktop\RB\Rocketbot\modules\ExcelOnline\docs\imgs\Banner_ExcelOnline.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.  

## Como usar este modulo

Antes de usar este modulo, es necesario registrar tu aplicación en el portal de Azure. 

Registración: 

1. Inicie sesión en el portal de Azure. Luego diríjase al siguiente link: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade (Registración de Aplicaciones). 
2. Click en "Nueva Registración". Elija un nombre. 
3. En "Tipo de Cuentas" elija "Cuentas de cualquier directorio de la organización y cuentas personales de Microsoft (por ejemplo, Skype, Xbox)". 
4. En "URI de Redirección" selecciones "Web" como plataforma y ponga como URI: __https://localhost/__. Finalmente clieckear en "Registrar". 
5. Una vez registrada, en la sección "General" encontrara el "ID de Aplicación (Cliente)", escriba/guárdelo, lo necesitara más adelante. 
6. Diríjase a "Certificados y Secretos", genere un nuevo "Secreto de Cliente", escriba una descripción y fije la expiración en 24 meses (preferiblemente). Click en "Adherir" y escriba/guarde el "Valor" (NO el "ID de Secreto"), lo necesitara luego junto con el ID de la App. 
7. Finalmente, vaya a "Permisos de API", clickee en "Adherir permiso", luego en "Microsoft Graph" y seleccione "Permisos delegados". En la barra de búsqueda tipee, "Files.ReadWrite.All", marque con una tilde el casillero y clickee en "Adherir Permisos".
8. Obtener es código de acceso mediante la función __Obtener código de acceso__ utilizando el "ID de Aplicación (Cliente)" y "Secreto de Cliente" previamente guardados. Al aceptar, se abrirá una ventana del navegador web pidiendo que ingrese y acepte la conexión con la aplicación. Una vez de "Continuar", será redirigido a una URL como la siguiente:  https://localhost/?code=M.R3_BAY.c0173ccb-c865-d99f-008c-2c7c9478d63f. Copie lo que se encuentra luego de "code=", este es el código con el cual se solicitaran las credenciales de acceso mediante la función __Establecer credenciales__.
9. Utilice el código para generar por primera vez las credenciales mediante __Establecer credenciales__. Una vez generadas, ya no sera necesario realizar nuevamente los pasos anteriores hasta tanto expire el Secreto de Cliente (obtenido en el punto 6).

## __Descripción de los comandos__

### __Obtener codigo de acceso__
Obtener codigo de acceso para generar las credenciales de la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Cliente|ID de Cliente de la API|id_cliente|
|Valor Secreto Cliente|Secreto del Cliente (Valor) de la API|valor_secreto_cliente|

### __Establecer credenciales__
Establece las credenciales para tener disponible la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|auth_code|Código de autorización|codigo|
|Variable a asignar|Variable a asignar. Si la conexion es exitosa retornara True, caso contraria sera False|variable|

### __Obtener archivos XLSX__
Devuelve una lista con todos los archivos XLSX en la carpeta
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Variable a asignar|Variable a asignar. Retorna la lista de archivos|lista_archivos|

### __Obtener hojas de trabajo__
Obtiene hojas de trabajo de un archivo de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del Libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Variable a asignar|Variable a asignar. Retorna la lista de hojas del libro|lista_hojas|

### __Crear libro__
Crear un nuevo libro en la ubicación por defecto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre del libro|Nombre del libro|Libro Nuevo|
|Variable a asignar|Variable a asignar. Retorna el ID del nuevo libro creado|id_nuevoLibro|

### __Añadir nueva hoja__
Añadir una nueva hoja al libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Variable a asignar|Variable a asignar. Retorna el nombre de la nueva hoja|nombre_nuevaHoja|

### __Obtener celda__
Obtener valor de celda o rango
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del Libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Celda o rango|Celda o rango|A1:B1|
|Variable a asignar|Variable a asignar. Retorna el valor de la celda o rango|valor_celda|

### __Escribir/cambiar celda__
Escribir/cambiar el valor de una celda o rango
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID del Libro|ID del libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Celda o rango|Celda o rango donde se escribira|A1:B1|
|Valor celda o rango|Variable a asignar. Retorna nuevos valores de la celda o rango"|nuevo_valor_celda|
