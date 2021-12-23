## Como usar este modulo

Para permitir la autenticación, primero debe registrar su aplicación en Azure App Registrations.

1- Inicie sesión en Azure Portal (App Registrations)
2- Cree una aplicación. Establezca un nombre.
3- En “Tipos de cuenta” soportados elige "Cuentas de cualquier directorio de la organización y cuentas personales de Microsoft (por ejemplo, Skype, Xbox, Outlook.com)".
4- Establezca la uri de redirección (Web) : https://localhost y haga click en registrarse.
5- Anote el ID de la aplicación (cliente). Necesitará este valor.
6- En "Certificados y secretos", genere un nuevo secreto de cliente. Establezca que la caducidad sea preferiblemente 24 meses. Anote el VALOR del secreto de cliente creado ahora. Se ocultará más adelante. Debe copiar VALOR, no Id de secreto.
7- En Permisos de API, añadir los siguientes permisos, en permisos delegados y de aplicación: Files.ReadWriteAll


---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


## How to use this module

To enable authentication, you must first register your application in Azure App Registrations.

1- Log in to Azure Portal (App Registrations)
2- Create an application. Set a name.
3- Under "Supported account types" choose "Accounts from any organization directory and Microsoft personal accounts (e.g. Skype, Xbox, Outlook.com)".
4- Set the redirection uri (Web) : https://localhost and click on register.
5- Write down the application (client) ID. You will need this value.
6- In "Certificates and secrets", generate a new client secret. Set the expiration date preferably to 24 months. Write down the VALUE of the client secret created now. It will be hidden later. You must copy VALUE, not secret ID.
7- In API Permissions, add the following permissions, in delegated and application permissions: Files.ReadWriteAll

