



# ExcelOnline
  
Working with Excel Online  
  
![banner](/docs/imgs/Banner_C:\Users\jmsir\Desktop\RB\Rocketbot\modules\ExcelOnline.png)
## Como instalar este módulo
  
__Baixe__ e __instale__ o conteúdo na pasta 'modules' no caminho do Rocketbot  



## Descrição do comando

### Definir credenciais
  
Defina as credenciais para ter a API disponível
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|client_id|Client ID da API|Your client_id|
|client_secret|Client Secret da API|Your client_secret|
|redirect_uri|Redirecionamento da API|http://localhost:5000|
|code|Código de autorização|code|
|tenant|Tenant da API|tenant|
|conexão|Variável onde armazenaremos nosso resultado. Se a conexão for bem sucedida retornará True, caso contrário será False|connection|

### Obter código de acesso
  
Obtenha o código de acesso para gerar credenciais de API
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|client_id|Client ID da API|cliente_id|
|valor_segredo_cliente|Valor secreto do cliente da API|valor_secreto_cliente|

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
|Variável a atribuir|Variável a atribuir. Retorna a lista de arquivos|lista_archivos|

### Obter planilhas
  
Obter planilhas de um arquivo do Excel
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Variável a atribuir|Variable a asignar. Retorna la lista de hojas del libro|lista_hojas|

### Criar livro
  
Criar um novo livro no local padrão
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome do livro|Nome do livro|Libro Nuevo|
|Variable a asignar|Variável a atribuir. ID de retorno do novo livro|id_nuevoLibro|

### Adicionar nova planilha
  
Adicionar uma nova folha ao livro
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nome da planilha|Nome da planilha|Sheet1|
|Variável a atribuir|Variável a atribuir. Retorna o nome da nova planilha|nombre_nuevaHoja|

### Obter célula
  
Obter valor da célula ou intervalo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nome da planilha|Nome da planilha|Sheet1|
|Célula ou intervalo|Célula ou intervalo|A1:B1|
|Variável a atribuir|Variável a atribuir. Retorna o valor da célula ou intervalo|valor_celda|

### Gravar/alterar célula
  
Digite/altere o valor de uma célula ou intervalo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|ID do livro|ID do livro|FB60B3125CDC0C03!238 (20 digits ID Code)|
|Nome da planilha|Nome da planilha|Sheet1|
|Célula ou intervalo|Célula ou intervalo onde será escrito|A1:B1|
|Valor ou intervalo da célula|Valor ou intervalo da célula|value_cell|
|Variável a atribuir|Variável a atribuir. Retorna True se a alteração foi bem sucedida, caso contrário será False|nuevo_valor_celda|
