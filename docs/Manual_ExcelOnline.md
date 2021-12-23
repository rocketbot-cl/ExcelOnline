



# ExcelOnline
  
Working with Excel Online  
  
![banner](https://i.imgur.com/A3LR8l1.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.  



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
|res|Variable donde guardaremos nuestro resultado. Si la conexion es exitosa retornara True, caso contraria sera False|res|

### Obtener archivos XLSX
  
Devuelve una lista con todos los archivos XLSX en la carpeta
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Variable a asignar|Variable a asignar. Retorna la lista de archivos|lista_archivos|

### Obtener hojas de trabajo
  
Obtiene hojas de trabajo de un archivo de excel
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID de Libro de trabajo|ID de Libro de trabajo|FB60B3125CDC0C03!239|
|Variable a asignar|Variable a asignar. Retorna la lista de hojas del libro|lista_hojas|

### Crear libro
  
Crear un nuevo libro en la ubicación por defecto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre del libro|Nombre del libro|Libro Nuevo|
|Variable a asignar|Variable a asignar. Retorna los datos del nuevo libro creado|datos_nuevolibro|

### Añadir nueva hoja
  
Añadir una nueva hoja al libro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Libro|ID del libro|FB60B3125CDC0C03!238|
|Nombre de la hoja|Nombre de la hoja|Hoja 2|

### Obtener celda
  
Obtener valor de celda
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Libro|ID del libro|FB60B3125CDC0C03!238|
|Nombre de la hoja|Nombre de la hoja|Hoja 1|
|Celda o rango|Celda o rango|A1:B1|
|Variable a asignar|Variable a asignar. Retorna el valor de la celda o rango|value_cell|

### Escribir celda
  
Escribir una celda en una hoja
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|ID Libro|ID del libro|FB60B3125CDC0C03!238|
|Nombre de la hoja|Nombre de la hoja|Hoja 1|
|Celda o rango|Celda o rango donde se escribira|A1:B1|
|Valor celda o rango|Valor de la celda o rango|value_cell|
