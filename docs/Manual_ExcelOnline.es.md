



# ExcelOnline
  
Cree, lea, actualice y elimine datos de archivos de Excel almacenados en OneDrive  

*Read this in other languages: [English](Manual_ExcelOnline.md), [Português](Manual_ExcelOnline.pr.md), [Español](Manual_ExcelOnline.es.md)*
  
![banner](imgs/Banner_ExcelOnline.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  



## Como usar este modulo

Antes de usar este módulo, es necesario registrar tu aplicación en el portal de Azure App Registrations. 

1. Inicie sesión en Azure Portal y busque el servicio Azure Active Directory.
2. En el menu en el lateral izquierdo, ingrese a "Registros de Aplicaciones".
3. Seleccione "Nuevo registro".
4. En “Tipos de cuenta compatibles” soportados elija:
    - "Cuentas en cualquier directorio organizativo (cualquier directorio de Azure AD: multiinquilino) y cuentas de Microsoft personales (como Skype o Xbox)" para este caso utilizar  ID Inquilino = **common**.
    - "Solo cuentas de este directorio organizativo (solo esta cuenta: inquilino único) para este caso utilizar **ID Inquilino** especifico de la aplicación.
    - "Solo cuentas personales de Microsoft " for this case use use Tenant ID = **consumers**.
5. En "URI de Redirección" selecciones "Web" como plataforma y ponga como URI: __https://localhost/__. Finalmente clieckear en "Registrar". 
6. Copie el ID de la aplicación (cliente). Necesitará este valor. 
7. Dentro de "Certificados y secretos", genere un nuevo secreto de cliente. Establezca la caducidad (preferiblemente 24 meses). Copie el **VALOR** del secreto de cliente creado (**__NO el ID de Secreto__**). El mismo se ocultará al cabo de unos minutos.
8. Finalmente, vaya a "Permisos de API", haga click en "Agregar un permiso", seleccione "Microsoft Graph", luego "Permisos delegados". En la barra de búsqueda tipee, "Files.ReadWriteAll", marque con un tilde el casillero y clickee en "Adherir Permisos". 


## Descripción de los comandos

### Obtener código de acceso
  
Obtener código de acceso para generar las credenciales de la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|id_cliente|ID de Cliente de la API|cliente_id|
|valor_secreto_cliente|Valor del Secreto del Cliente de la API|valor_secreto_cliente|
|tenant|Tenant ID de la API|common|

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
|ID de Disco Compartido (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|Variable a asignar|Variable a asignar. Retorna la lista de archivos|lista_archivos|

### Obtener hojas de trabajo
  
Obtiene hojas de trabajo de un archivo de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Disco Compartido (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID del Libro|ID del Libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Variable a asignar|Variable a asignar. Retorna la lista de hojas del libro|lista_hojas|

### Crear libro
  
Crear un nuevo libro en la ubicación por defecto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Disco Compartido (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|Nombre del libro|Nombre del libro|Libro Nuevo|
|Variable a asignar|Variable a asignar. Retorna ID del nuevo libro|id_nuevoLibro|

### Añadir nueva hoja
  
Añadir una nueva hoja al libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Disco Compartido (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID del Libro|ID del libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Variable a asignar|Variable a asignar. Retorna el nombre de la nueva hoja|nombre_nuevaHoja|

### Obtener celda
  
Obtener valor de celda o rango
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Disco Compartido (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID del Libro|ID del Libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Celda o rango|Celda o rango A1B2|Cell or Range|
|Variable a asignar|Variable a asignar. Retorna el valor de la celda o rango|valor_celda|

### Escribir/cambiar celda
  
Escribir/cambiar el valor de una celda o rango
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Disco Compartido (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID del Libro|ID del libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Celda o rango|Celda o rango donde se escribira A1B2|Cell or Range|
|Valor celda o rango|Valor de la celda o rango [[1,2],[1,2]]|Value or list of values|
|Variable a asignar|Variable a asignar. Retorna True si el cambio fue exitoso, caso contraria sera False|nuevo_valor_celda|

### Escribir formula
  
Escribir una fórmula en una celda o rango
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Disco Compartido (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID del Libro|ID del libro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nombre de la hoja|Nombre de la hoja|Sheet1|
|Celda o rango|Celda o rango donde se escribira A1B2|Cell or Range|
|Formula|Formula en formato Excel|=sum(2;2)|
|Variable a asignar|Variable a asignar. Retorna True si el cambio fue exitoso, caso contraria sera False|nuevo_valor_celda|

### (OBSOLETO) Establecer credenciales
  
Establece las credenciales para tener disponible la API
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|client_id|Client ID de la API|Your client_id|
|client_secret|Client Secret de la API|Your client_secret|
|redirect_uri|Redirección de la API|http://localhost:5000|
|code|Código de autorización|code|
|tenant|Tenant de la API|tenant|
|conexion|Variable donde guardaremos nuestro resultado. Si la conexion es exitosa retornara True, caso contraria sera False|connection|
