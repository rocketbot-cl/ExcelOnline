


## How to use this module

Before using this module, you need to register your app in the Azure App Registrations portal.

1. Sign in to the Azure portal and search for the Azure Active Directory service.
2. On the left side menu, get into "App Registrations".
3. Select "New record".
4. Under “Compatible account types” supported choose:
    - "Accounts in any organizational directory (any Azure AD directory: multi-tenant) and personal Microsoft accounts (such as Skype or Xbox)" for this case use Tenant ID = **common**.
    - "Only accounts from this organizational directory (only this account: single tenant) for this case use application-specific **Tenant ID**.
    - "Personal Microsoft accounts only" for this case use use Tenant ID = **consumers**.
5. In "Redirect URI" select "Web" as a plataform and set the URI to: https://localhost/. Finally click on "Register"
6. Copy the application (client) ID. You will need this value.
7. Under "Certificates and secrets", generate a new client secret. Set the expiration (preferably 24 months). Copy the **VALUE** of the created client secret (**__NOT the Secret ID__**). It will hide after a few minutes.
8. Finally, go to "API persmissions", click "Add a permission", select "Microsoft Graph", then "Delegated permissions". In the search bar type "Files.ReadWriteAll", mark the checkbox and click on "Add permissions".

---

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

---

## Como usar este módulo

Antes de usar este módulo, você precisa registrar seu aplicativo no portal de Registros de Aplicativo do Azure.

1. Entre no portal do Azure e procure o serviço Azure Active Directory.
2. No menu do lado esquerdo, entre em "Registros de aplicativos".
3. Selecione "Novo registro".
4. Em "Tipos de conta compatíveis" suportados, escolha:
    - "Contas em qualquer diretório organizacional (qualquer diretório do Azure AD: multilocatário) e contas pessoais da Microsoft (como Skype ou Xbox)" para este caso, use ID de locatário = **common**.
    - "Somente contas deste diretório organizacional (somente esta conta: locatário único) para este caso usam **ID de locatário específico** do aplicativo.
    - "Somente contas pessoais da Microsoft" para este caso, use ID do locatário = **consumidores**.
5. Em "Redirect URI" selecione "Web" como plataforma e defina o URI para: https://localhost/. Por fim, clique em "Registrar"
6. Copie o ID do aplicativo (cliente). Você vai precisar desse valor.
7. Em "Certificados e segredos", gere um novo segredo do cliente. Defina a validade (de preferência 24 meses). Copie o **VALUE** do segredo do cliente criado (**__NÃO o ID do segredo__**). Ele vai esconder depois de alguns minutos.
8. Por fim, vá em "Permissões de API", clique em "Adicionar uma permissão", selecione "Microsoft Graph", depois em "Permissões delegadas". Na barra de pesquisa digite "Files.ReadWriteAll", marque a caixa de seleção e clique em "Adicionar permissões".
