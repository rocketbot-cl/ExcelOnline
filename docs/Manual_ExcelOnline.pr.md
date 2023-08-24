



# ExcelOnline
  
Crie, leia, atualize e exclua dados de arquivos do Excel armazenados no OneDrive  

*Read this in other languages: [English](Manual_ExcelOnline.md), [Português](Manual_ExcelOnline.pr.md), [Español](Manual_ExcelOnline.es.md)*
  
![banner](imgs/Banner_ExcelOnline.png)
## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  

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


## Descrição do comando

### Obter código de acesso
  
Obtenha o código de acesso para gerar credenciais de API
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|client_id|Client ID da API|cliente_id|
|valor_segredo_cliente|Valor secreto do cliente da API|valor_secreto_cliente|
|tenant|Tenant ID da API|common|

### Definir credenciais
  
Defina as credenciais para ter a API disponível
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|code|Código de autorização|code|
|Variável a atribuir|Variável a atribuir. Se a conexão for bem sucedida retorna True, caso contrário será False|variable|

### Obter arquivos XLSX
  
Retorna uma lista de todos os arquivos XLSX no local padrão
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do disco compartilhado (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|Variável a atribuir|Variável a atribuir. Retorna a lista de arquivos|lista_archivos|

### Obter planilhas
  
Obter planilhas de um arquivo do Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do disco compartilhado (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Variável a atribuir|Variable a asignar. Retorna la lista de hojas del libro|lista_hojas|

### Criar livro
  
Criar um novo livro no local padrão
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do disco compartilhado (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|Nome do livro|Nome do livro|Libro Nuevo|
|Variable a asignar|Variável a atribuir. ID de retorno do novo livro|id_nuevoLibro|

### Adicionar nova planilha
  
Adicionar uma nova folha ao livro
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do disco compartilhado (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nome da planilha|Nome da planilha|Sheet1|
|Variável a atribuir|Variável a atribuir. Retorna o nome da nova planilha|nombre_nuevaHoja|

### Obter célula
  
Obter valor da célula ou intervalo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do disco compartilhado (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nome da planilha|Nome da planilha|Sheet1|
|Célula ou intervalo|Célula ou intervalo A1B2|Cell or Range|
|Variável a atribuir|Variável a atribuir. Retorna o valor da célula ou intervalo|valor_celda|

### Gravar/alterar célula
  
Digite/altere o valor de uma célula ou intervalo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do disco compartilhado (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nome da planilha|Nome da planilha|Sheet1|
|Célula ou intervalo|Célula ou intervalo onde será escrito A1B2|Cell or Range|
|Valor ou intervalo da célula|Valor ou intervalo da célula [[1,2],[1,2]]|Value or list of values|
|Variável a atribuir|Variável a atribuir. Retorna True se a alteração foi bem sucedida, caso contrário será False|nuevo_valor_celda|

### Escrever fórmula
  
Escrever uma fórmula em uma célula ou intervalo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do disco compartilhado (Opcional)||b!4Zasr9LvqUiwt4OZ8irYdG3gm207yiJPkTu3c6KrXmFKVLpG3_FZTrGY-Gxn974J|
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nome da planilha|Nome da planilha|Sheet1|
|Célula ou intervalo|Célula ou intervalo onde será escrito A1B2|Cell or Range|
|Fórmula|Fórmula em formato Excel|=sum(2;2)|
|Variável a atribuir|Variável a atribuir. Retorna True se a alteração foi bem sucedida, caso contrário será False|nuevo_valor_celda|

### (DESCONTINUADO) Definir credenciais
  
Defina as credenciais para ter a API disponível
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|client_id|Client ID da API|Your client_id|
|client_secret|Client Secret da API|Your client_secret|
|redirect_uri|Redirecionamento da API|http://localhost:5000|
|code|Código de autorização|code|
|tenant|Tenant da API|tenant|
|conexão|Variável onde armazenaremos nosso resultado. Se a conexão for bem sucedida retornará True, caso contrário será False|connection|
