



# ExcelOnline
  
Trabalhando com o Excel Online  

## Como instalar este módulo
  
__Baixe__ e __instale__ o conteúdo na pasta 'modules' no caminho do Rocketbot  

## Como usar este módulo

Antes de usar este módulo, você deve registrar seu aplicativo no Portal do Azure.

Registros.

1. Entre no Portal do Azure. Em seguida, vá para o próximo link: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade (Registros de aplicativos).
2. Clique em "Novo registro". Escolha um nome.
3. Em "Tipos de conta compatíveis", escolha "Contas em qualquer diretório organizacional (qualquer diretório do Azure AD - Multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)".
4. Em "Redirect URI" selecione "Web" como plataforma e defina o URI para: https://localhost/. Por fim, clique em "Registrar"
5. Uma vez cadastrado, na seção "Visão geral" você encontrará o "ID do aplicativo (cliente)", anote/salve-o, você precisará dele mais tarde.
6. Vá em "Certificados e segredos", gere um "Novo segredo de cliente", escreva uma descrição e defina o tempo de expração para 24 meses (de preferência). Clique em "Add" e anote/salve o "Value" (NÃO o "Secret ID"), você precisará dele mais tarde junto com o App ID.
7. Por fim, vá em "Permissões de API", clique em "Adicionar uma permissão", depois em "Microsoft Graph" e selecione "Permissões delegadas". Na barra de pesquisa digite "Files.ReadWriteAll", marque a caixa de seleção e clique em "Adicionar permissões".

## Overview


1. Definir credenciais  
Defina as credenciais para ter a API disponível

2. Obter código de acesso  
Obtenha o código de acesso para gerar credenciais de API

3. Definir credenciais  
Defina as credenciais para ter a API disponível

4. Obter arquivos XLSX  
Retorna uma lista de todos os arquivos XLSX no local padrão

5. Obter planilhas  
Obter planilhas de um arquivo do Excel

6. Criar livro  
Criar um novo livro no local padrão

7. Adicionar nova planilha  
Adicionar uma nova folha ao livro

8. Obter célula  
Obter valor da célula ou intervalo

9. Gravar/alterar célula  
Digite/altere o valor de uma célula ou intervalo

----
### OS

- windows
- mac
- linux

### Dependencies

### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)