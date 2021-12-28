## Como usar este modulo

Para permitir la autenticación, primero debe registrar su aplicación en Azure App Registrations.

1- Inicie sesión en Azure Portal (App Registrations). Link aqui: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
2- Cree una aplicación. Establezca un nombre.
3- En “Tipos de cuenta” soportados elige "Cuentas de cualquier directorio de la organización y cuentas personales de Microsoft (por ejemplo, Skype, Xbox, Outlook.com)".
4- Establezca la uri de redirección como https://localhost. Establezca la plataforma como Web y haga click en registrarse.
5- Anote el ID de la aplicación (cliente). Necesitará este valor.
6- En "Certificados y secretos", genere un nuevo secreto de cliente. Establezca que la caducidad sea preferiblemente 24 meses. Anote el VALOR del secreto de cliente creado ahora. Se ocultará más adelante. Debe copiar VALOR, no Id de secreto.
7- En Permisos de API, damos clic en "Añadir Permisos", luego elegimos la opcion de "Microsoft Graph" y luego añadir los siguientes permisos en permisos delegados y de aplicación: Files.ReadWriteAll
8- Luego modificar los valores entre llaves de el el siguiente link con los datos de su aplicacion: https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&response_mode=query&scope=offline_access%20files.readwrite.all&state=12345
Deberia modificar tenant por el que le aparece en "Informacion general" en "Id de directorio (Inquilino). En caso de tener una cuenta personal, utilizar "common" como tenant. Reemplazar el client_id y redirect_uri correspondiente. Nos pedira que iniciemos sesion y demos permisos. Luego al final, se generara una URL en nuestra barra de direcciones como la siguiente: http://localhost/?code=M.R3_BAY.3a97b36e-73cc-2342-3421-0b1054c192d6&state=12345 y obtener el CODE que en este caso seria el siguiente: M.R3_BAY.3a97b36e-73cc-2342-3421-0b1054c192d6

---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


## How to use this module

To enable authentication, you must first register your application in Azure App Registrations.

1- Log in to Azure Portal (App Registrations). Link here: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
2- Create an application. Set a name.
3- Under "Supported account types" choose "Accounts from any organization directory and Microsoft personal accounts (e.g. Skype, Xbox, Outlook.com)".
4- Set the redirect uri as https://localhost. Set the platform as Web and click on register.
5- Write down the application (client) ID. You will need this value.
6- In "Certificates and secrets", generate a new client secret. Set the expiration date preferably to 24 months. Write down the VALUE of the client secret created now. It will be hidden later. You must copy VALUE, not secret ID.
7- In API Permissions, click on "Add Permissions", then choose the "Microsoft Graph" option and then add the following permissions in delegated and application permissions: Files.ReadWriteAll
8- Then modify the values between braces of the following link with the data of your application: https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id={client_id}&response_type=code&redirect_uri={redirect_uri}&response_mode=query&scope=offline_access%20files.readwrite.all&state=12345
You should change tenant to the one that appears in "General Information" under "Directory ID (Tenant). In case you have a personal account, use "common" as tenant. Replace the corresponding client_id and redirect_uri. We will be asked to log in and give permissions. Then at the end, it will generate a URL in our address bar like the following: http://localhost/?code=M.R3_BAY.3a97b36e-73cc-2342-3421-0b1054c192d6&state=12345 and get the CODE which in this case would be the following: M.R3_BAY.3a97b36e-73cc-2342-3421-0b1054c192d6


