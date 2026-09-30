# CLEAN — Análise de repositórios de arquivos e caixas de e-mail

Aplicação web em Node.js que percorre repositórios de arquivos do Windows (pastas locais e
compartilhamentos de rede) e do **Microsoft 365 (OneDrive e SharePoint)**, procura os termos de uma
**lista de referência** no **nome** e no **conteúdo** dos arquivos e gera relatórios com **a
informação encontrada** e **o último usuário que interagiu com cada arquivo**. A seção **E-mail** faz a mesma busca em **caixas de e-mail
pré-configuradas** (Microsoft 365, Google Workspace ou servidores IMAP): varre todas as caixas e
pastas e procura os termos no assunto, no **corpo** e nos **anexos** de cada mensagem.

- Listas de referência com termos de texto (sem diferenciar maiúsculas e acentos: "salario" encontra
  "SALÁRIO") ou expressões regulares, com modelos prontos e validados: CPF, CNPJ (inclusive o novo
  CNPJ alfanumérico), PIS/PASEP, e-mail, telefone, cartão de crédito e "senha em texto".
- Lê o conteúdo de Word, Excel e PowerPoint (novos e 97-2003), PDF, OpenDocument (LibreOffice),
  RTF, e-mails (.msg e .eml), páginas salvas (.mht), HTML, textos e CSV (UTF-8, UTF-16 e
  Windows-1252) e os nomes dos arquivos dentro de .zip. Arquivos protegidos por senha são
  identificados.
- Último usuário a partir de três fontes, da mais precisa para a menos precisa: **log de auditoria
  do Windows** (quem acessou ou alterou por último), **metadados do documento** ("salvo por último
  por") e **proprietário do arquivo (NTFS)**. No OneDrive e no SharePoint, quem alterou o arquivo
  por último segundo o Microsoft 365.
- **OneDrive** (todas as contas do locatário ou as escolhidas) e **SharePoint** (todos os sites ou
  os escolhidos, com os subsites e todas as bibliotecas de documentos), pela API Microsoft Graph,
  com a mesma busca no nome e no conteúdo, o mesmo relatório e a mesma exclusão opcional.
- Relatório na tela com filtros, gráficos e trechos em que cada termo aparece (com página, planilha,
  slide ou linha), e exportação para **Excel**, **CSV**, **HTML** (para imprimir) e **JSON**.
- Caixas de e-mail do **Microsoft 365** (API Microsoft Graph), do **Google Workspace** (API do
  Gmail) ou de qualquer servidor **IMAP**: todas as caixas do locatário/domínio ou uma lista, todas
  as pastas, assunto, corpo, nomes e conteúdo dos anexos (os mesmos formatos acima, inclusive
  e-mails encaminhados como anexo). O relatório mostra a caixa, a pasta, o remetente, os
  destinatários e a data de cada mensagem encontrada. Senhas e chaves ficam gravadas cifradas.
- **Raio-X das caixas**: a análise de e-mail também gera um retrato de todas as mensagens analisadas
  (não só as com ocorrências) — quantidade por pasta, e-mails mais antigos e a distribuição ao longo
  do tempo.
- **Listagens de e-mail** (somente leitura, sem procurar termos e sem excluir): o **catálogo de todas
  as contas registradas no domínio** (nome, endereço principal, apelidos, situação, tipo, licença,
  data de criação, departamento) e a **listagem de todas as mensagens de cada caixa** com os dados de
  cada uma (remetente, destinatários, data, assunto, pasta, tamanho) — com relatório, filtros,
  gráficos e exportação.
- **OAuth 2.0 da Microsoft** em todas as formas usadas pelo Exchange Online e pelo Outlook.com:
  aplicativo com **segredo do cliente**, aplicativo com **certificado** (gerado pelo próprio CLEAN)
  ou **conta Microsoft conectada** (a pessoa entra com a conta pelo código de dispositivo; a
  autorização é renovada sozinha) — também no **IMAP** do Exchange Online (XOAUTH2), que não aceita
  mais senha.
- **Busca por tipo de arquivo**: vídeos, músicas, imagens, executáveis, compactados, imagens de
  disco, temporários ou as extensões que você informar — com tamanho mínimo e a conferência do tipo
  real pelo conteúdo (acha um vídeo renomeado para .pdf), para revisão no relatório ou exclusão
  automática, agora ou agendada.
- **Exclusão** opcional dos arquivos e mensagens encontrados: automática durante a análise
  ("analisar e excluir"), item a item pelo relatório ou, nos relatórios de arquivos, **em lote** (os
  selecionados ou todos os filtrados), com registro de cada exclusão.
- **Agendamentos**: análises executadas sozinhas, uma vez ou com repetição (a cada algumas horas,
  diária, semanal ou mensal), com análise incremental, histórico e retenção dos relatórios.
- **Políticas de retenção**: eliminam os arquivos e as mensagens mais antigos que uma idade máxima
  (ex.: arquivos sem uso há mais de 5 anos, e-mails recebidos há mais de 10 anos), pela data da
  última modificação, do último acesso, da criação ou do recebimento — com simulação, limite de
  exclusões por execução e execução manual ou agendada.
- Sem banco de dados e sem etapa de compilação: basta instalar o Node.js e executar.

## Sumário

1. [Requisitos](#requisitos)
2. [Instalação e primeiro uso](#instalação-e-primeiro-uso)
3. [Como usar](#como-usar)
4. [Como o último usuário é identificado](#como-o-último-usuário-é-identificado)
5. [Habilitando a auditoria do Windows (opcional)](#habilitando-a-auditoria-do-windows-opcional)
6. [Conta de serviço e permissões](#conta-de-serviço-e-permissões)
7. [OneDrive e SharePoint](#onedrive-e-sharepoint)
8. [Análise de caixas de e-mail](#análise-de-caixas-de-e-mail)
   - [Microsoft 365 (Exchange Online)](#microsoft-365-exchange-online)
   - [Autenticação na Microsoft (OAuth 2.0)](#autenticação-na-microsoft-oauth-20)
   - [Google Workspace (Gmail)](#google-workspace-gmail)
   - [Servidores IMAP](#servidores-imap)
   - [Listagens de e-mail (contas e mensagens)](#listagens-de-e-mail-contas-e-mensagens)
9. [Exclusão dos itens encontrados](#exclusão-dos-itens-encontrados)
   - [Exclusão em lote pelo relatório](#exclusão-em-lote-pelo-relatório)
10. [Agendamentos](#agendamentos)
11. [Políticas de retenção](#políticas-de-retenção)
12. [Busca por tipo de arquivo](#busca-por-tipo-de-arquivo)
13. [Executando como serviço](#executando-como-serviço)
14. [Configuração](#configuração)
15. [Segurança](#segurança)
16. [Formatos suportados e limitações](#formatos-suportados-e-limitações)
17. [Desenvolvimento](#desenvolvimento)

## Requisitos

- Windows 10/11 ou Windows Server 2016 ou superior (também funciona em Linux, para pastas locais ou
  compartilhamentos montados; nesse caso o proprietário vem do UID e não há log de auditoria).
- [Node.js](https://nodejs.org) **22.13 ou superior** (versão LTS recomendada).
- Windows PowerShell 5.1 (já incluso no Windows) — usado para ler o proprietário dos arquivos e o log
  de auditoria.

## Instalação e primeiro uso

1. Copie a pasta do projeto para o servidor, por exemplo `C:\CLEAN`.
2. Dê dois cliques em **`iniciar.bat`** (na primeira vez ele instala as dependências com `npm install`).
   Pelo terminal, o equivalente é:

   ```bat
   cd C:\CLEAN
   npm install --omit=dev
   npm start
   ```

3. Abra **http://localhost:3000** no navegador.

Para experimentar com arquivos de exemplo, rode `npm run demo`: ele cria a pasta
`demo\Compartilhamento` com documentos de teste e cadastra um repositório e uma lista de referência
de exemplo. Depois é só iniciar uma análise.

## Como usar

1. **Repositórios** — cadastre as pastas a analisar: `D:\Dados\RH` ou `\\servidor\compartilhamento`.
   O botão *Testar acesso* confirma se a conta do CLEAN consegue ler a pasta. É possível ignorar
   arquivos e pastas por padrões (`*.bak`, `Backup`, `Financeiro\Antigo`); `$RECYCLE.BIN`,
   `System Volume Information`, arquivos temporários do Office (`~$*`) e `Thumbs.db` já são ignorados.
   Contas do **OneDrive** e sites do **SharePoint** também podem ser cadastrados como repositórios
   (veja [OneDrive e SharePoint](#onedrive-e-sharepoint)).
2. **Listas de referência** — crie uma lista e adicione termos:
   - **Texto**: ignora maiúsculas, acentos e espaços repetidos. Marque *Palavra inteira* para que
     "ana" não encontre "banana".
   - **Expressão regular** (sintaxe JavaScript), opcionalmente com validação (CPF, CNPJ, PIS,
     cartão) para descartar números inválidos, e um nome para o relatório.
   - *Adicionar vários*, *Importar .txt/.csv* (um termo por linha ou a primeira coluna do CSV) e
     *Adicionar modelo pronto*.
   - *Testar os termos* mostra o que seria encontrado em um texto de exemplo antes de salvar.
3. **Nova análise** — em *O que procurar*, escolha **Termos das listas de referência** ou **Tipos de
   arquivo** (veja [Busca por tipo de arquivo](#busca-por-tipo-de-arquivo)). Nos termos, escolha
   repositórios e listas e o que verificar:
   - nome do arquivo (ou o caminho completo, incluindo os nomes das pastas) e/ou conteúdo;
   - somente arquivos alterados a partir de uma data (vale a data mais recente entre a modificação
     e a criação: um arquivo copiado para a pasta depois da data entra na análise, mesmo mantendo a
     data de modificação original; mudanças só de permissões ou atributos não contam);
   - tamanho máximo para ler o conteúdo (arquivos maiores têm apenas o nome verificado; textos
     longos, só o início);
   - quantidade de arquivos processados em paralelo.
4. **Relatório** — acompanha o progresso em tempo real. Mostra os números da análise, os termos
   mais encontrados, os últimos usuários (clique numa barra para filtrar) e a lista de arquivos com
   ocorrências. Cada arquivo abre os detalhes: caminho, datas, todas as fontes de "último usuário" e
   os trechos com o termo destacado. As exportações respeitam os filtros aplicados:
   - **Excel**: abas *Resumo*, *Arquivos* (um arquivo por linha), *Ocorrências* (um termo por
     linha, com valores e trechos) e *Erros*;
   - **CSV**: uma linha por arquivo, termo e local (nome ou conteúdo), no padrão do Excel em
     português (`;` e UTF-8);
   - **HTML**: relatório para leitura ou impressão; **JSON**: para integração com outros sistemas.

As análises rodam em segundo plano (é possível fechar o navegador) e podem ser canceladas a qualquer
momento; o que já foi encontrado é mantido. Para repetir uma análise automaticamente (toda noite,
toda semana...), use **Quando executar › Agendar** ou o menu **Automação › Agendamentos** — veja
[Agendamentos](#agendamentos). Para eliminar os arquivos e os e-mails antigos (por exemplo, sem uso
há mais de 5 anos), use **Automação › Retenção** — veja [Políticas de retenção](#políticas-de-retenção).
Para achar (e, se quiser, excluir) arquivos de um tipo — vídeos, músicas, instaladores, backups —,
use **Buscar por tipo de arquivo** — veja [Busca por tipo de arquivo](#busca-por-tipo-de-arquivo).

Para as caixas de e-mail, use o grupo **E-mail** do menu: *Caixas de e-mail* (conexões) e *Análises
de e-mail* — veja [Análise de caixas de e-mail](#análise-de-caixas-de-e-mail).

## Como o último usuário é identificado

O NTFS não registra "quem alterou por último". O CLEAN combina as fontes disponíveis e informa no
relatório qual foi usada:

| Prioridade | Fonte | O que informa | Observações |
|---|---|---|---|
| 1 | **Log de auditoria** (eventos 4663 e 5145) | Quem acessou por último (e a ação: leitura, gravação, exclusão, permissões) e quem alterou por último | É a fonte mais fiel, mas exige a auditoria habilitada **antes** (ver abaixo) e vale apenas para o período mantido no log. |
| 2 | **Metadados do documento** | "Salvo por último por" do Word, Excel, PowerPoint, LibreOffice, RTF e .msg | É o nome configurado no Office de quem salvou (normalmente o nome completo). Não muda quando o arquivo é alterado por outros programas. |
| 3 | **Proprietário (NTFS)** | Dono do arquivo | Em geral quem **criou** o arquivo (às vezes o grupo Administradores). Usado quando não há as fontes acima. |

No **OneDrive** e no **SharePoint**, o último usuário é **quem alterou o arquivo por último**
segundo o Microsoft 365 (o relatório também mostra quem o criou e, no OneDrive, o dono da conta).

O relatório também mostra todas as fontes lado a lado (proprietário, autor, salvo por último por,
último acesso e última alteração na auditoria) para comparação.

A leitura feita pelo próprio CLEAN também gera eventos de auditoria; por isso a conta que executa o
CLEAN é sempre ignorada. Informe no repositório outras contas que leem todos os arquivos (backup,
antivírus, indexação) para que não apareçam como "último usuário".

## Habilitando a auditoria do Windows (opcional)

Faça no **servidor de arquivos** (via GPO ou localmente, como administrador).

1. **Política de auditoria** — em *Configuração do Computador › Políticas › Configurações do Windows ›
   Configurações de Segurança › Configuração Avançada de Política de Auditoria › Acesso a Objetos*,
   habilite *Êxito* em **Auditoria do Sistema de Arquivos** (gera o 4663) e, se desejar, em
   **Auditoria Detalhada de Compartilhamento de Arquivos** (gera o 5145, mais volumoso). Pelo
   terminal (os GUIDs funcionam em qualquer idioma do Windows):

   ```bat
   auditpol /set /subcategory:"{0CCE921D-69AE-11D9-BED3-505054503030}" /success:enable
   auditpol /set /subcategory:"{0CCE9244-69AE-11D9-BED3-505054503030}" /success:enable
   ```

2. **SACL da pasta** (necessária para o 4663) — *Propriedades › Segurança › Avançadas › Auditoria ›
   Adicionar*: entidade **Todos**, tipo **Êxito**, permissões *Criar arquivos/gravar dados*, *Criar
   pastas/acrescentar dados*, *Excluir* e, para registrar também leituras, *Listar pasta/ler dados*.
   Em PowerShell (como administrador):

   ```powershell
   $pasta = 'E:\Compartilhamentos\Financeiro'
   $acl = Get-Acl -Path $pasta -Audit
   $todos = New-Object System.Security.Principal.SecurityIdentifier('S-1-1-0')
   $regra = New-Object System.Security.AccessControl.FileSystemAuditRule($todos, 'ReadData,WriteData,AppendData,Delete', 'ContainerInherit,ObjectInherit', 'None', 'Success')
   $acl.AddAuditRule($regra)
   Set-Acl -Path $pasta -AclObject $acl
   ```

3. **Tamanho do log de Segurança** — aumente para que os eventos não sejam sobrescritos rapidamente
   (exemplo: 4 GB): `wevtutil sl Security /ms:4294967296`.

4. No CLEAN, edite o repositório e marque **Consultar o log de Segurança**. Para caminhos
   `\\servidor\compartilhamento`, o computador consultado é o servidor do caminho (pode ser alterado).
   O caminho local correspondente no servidor (por exemplo `E:\Compartilhamentos\Financeiro`) é
   descoberto pelos eventos 5145; se usar apenas o 4663, informe-o no campo *Caminho local no servidor*.
   Compartilhamentos administrativos (`\\servidor\E$\...`) são mapeados automaticamente.

## Conta de serviço e permissões

O CLEAN lê os arquivos com a conta do Windows que o executa. Recomenda-se uma conta de domínio
dedicada (por exemplo `EMPRESA\svc-clean`) com:

- permissão de **leitura** (NTFS e compartilhamento) nas pastas analisadas — e de **modificação**
  só nas pastas em que a exclusão for usada (veja *Exclusão dos itens encontrados*);
- participação no grupo **Leitores de Log de Eventos** (*Event Log Readers*) de cada servidor de
  arquivos, se usar a auditoria; para consultar outro computador, a regra de firewall
  *Gerenciamento Remoto do Log de Eventos* deve estar habilitada nele.

Pastas que a conta não consegue abrir aparecem na aba **Erros** do relatório.

## OneDrive e SharePoint

Em *Repositórios › Novo repositório*, escolha o tipo **OneDrive** ou **SharePoint**. A análise, o
relatório e a exclusão funcionam como nas pastas do Windows; os arquivos são lidos pela API
Microsoft Graph, com as permissões de um **registro de aplicativo** (sem usuário conectado).

- **OneDrive**: todas as contas do locatário (usuários sem OneDrive — nunca acessado, sem licença,
  salas, caixas compartilhadas — são ignorados e contados no relatório) ou somente as contas
  informadas (e-mail ou nome de logon do usuário). É possível ignorar contas por padrão (ex.:
  `teste@*`), pelo e-mail, pelo nome de logon ou pelo nome.
- **SharePoint**: todos os sites (sem os OneDrive pessoais, que são analisados pelo tipo OneDrive)
  ou somente os sites informados (endereço, ex.: `https://empresa.sharepoint.com/sites/Financeiro`;
  pode colar o endereço de uma página ou biblioteca do site). Os **subsites** e **todas as
  bibliotecas de documentos** de cada site são analisados — inclusive os arquivos das equipes do
  **Microsoft Teams**, que ficam nos sites delas. É possível ignorar sites por endereço (ignora
  também os subsites) ou por nome. Um endereço inexistente é informado como erro (o CLEAN não troca
  um subsite digitado errado pelo site acima dele); OneDrive pessoais não entram na lista de sites.
- Os padrões de *Ignorar* valem para pastas e arquivos e também para o nome das bibliotecas (ex.:
  `Site Assets`, `Style Library`); atalhos para pastas de outras bibliotecas ("Adicionar atalho a
  Meus arquivos") não são seguidos, para não analisar o mesmo arquivo duas vezes.
- O conteúdo é baixado para a memória do servidor do CLEAN até o limite de tamanho da análise
  (arquivos maiores têm só o nome verificado; textos longos, só o início) e lido pelos mesmos
  leitores dos arquivos do Windows. Nada é gravado em disco.
- O relatório mostra a conta ou o site, a biblioteca, a pasta, o endereço do arquivo (com link
  *Abrir no OneDrive/SharePoint*), quem alterou por último, quem criou e, no OneDrive, o dono.

### Registro do aplicativo no Microsoft Entra ID

Pode ser o mesmo das caixas de e-mail do Microsoft 365 — no formulário do repositório, escolha
*Usar as credenciais da conexão* (as credenciais ficam ligadas à conexão de e-mail: um novo segredo
ou certificado salvo nela passa a valer também para o repositório). Conexões com a **conta
conectada** não servem para o OneDrive e o SharePoint (as permissões dela são só de e-mail): o
repositório precisa das credenciais de um aplicativo, com segredo do cliente ou certificado.

Em *Registros de aplicativo › (o aplicativo) › Permissões de API › Adicionar › Microsoft Graph ›
Permissões de aplicativo*, inclua:

| Permissão | Para quê |
|---|---|
| `Files.Read.All` | Ler os arquivos de todos os OneDrive e sites. |
| `Sites.Read.All` | Listar os sites do SharePoint e as bibliotecas. |
| `User.Read.All` | Listar os usuários (OneDrive de todas as contas) e localizar as contas pelo e-mail. |
| `Files.ReadWrite.All` (no lugar de `Files.Read.All`) | Somente se for usar a exclusão. |

e clique em **Conceder consentimento do administrador**. Para limitar o CLEAN a alguns sites, use a
permissão `Sites.Selected` no lugar de `Sites.Read.All`/`Files.Read.All` e libere cada site para o
aplicativo (por exemplo, com o PnP PowerShell: `Grant-PnPEntraIDAppSitePermission -AppId <ID do
aplicativo> -DisplayName CLEAN -Site <endereço do site> -Permissions Read`, ou `Write` para excluir —
em versões mais antigas do PnP, `Grant-PnPAzureADAppSitePermission`); nesse caso, informe os sites
na lista (a opção "todos os sites" e o OneDrive precisam de `Sites.Read.All` e `Files.Read.All`).

*Testar conexão*, no formulário, confere as credenciais e o acesso a até três bibliotecas.

Limitações: só a versão atual de cada arquivo é analisada (versões anteriores e a lixeira, não);
blocos de anotações do OneNote não têm o conteúdo lido (só os nomes); a informação de **quem abriu
ou visualizou** um arquivo não é fornecida pelo Graph — para isso, use a pesquisa de auditoria do
Microsoft Purview.

## Análise de caixas de e-mail

A seção **E-mail** procura os termos das listas de referência nas mensagens de caixas de e-mail
cadastradas previamente:

1. **Caixas de e-mail** — cadastre uma ou mais conexões. Cada conexão é de um tipo:
   - **Microsoft 365** (Exchange Online; também contas pessoais do Outlook.com): pela API Microsoft
     Graph, com OAuth 2.0 — um registro de aplicativo (segredo do cliente ou certificado), que
     analisa **todas as caixas do locatário** (inclusive caixas compartilhadas) ou só as informadas,
     ou uma **conta Microsoft conectada**, que analisa a caixa dela e as compartilhadas com ela.
   - **Google Workspace** (Gmail): pela API do Gmail, com uma conta de serviço com delegação em todo
     o domínio. Analisa **todas as caixas do domínio** ou só as informadas.
   - **IMAP**: qualquer servidor IMAP (Exchange local, Zimbra, Dovecot, provedores de hospedagem...),
     com o login e a senha de cada caixa ou uma senha padrão de uma conta de serviço — ou, no
     Exchange Online e no Outlook.com, com o **OAuth 2.0 da Microsoft**.

   Em cada conexão é possível ignorar caixas (`noreply@*`) e pastas (`Pessoal`,
   `Caixa de Entrada/Newsletters`; aceita `*` e `?` e vale também para as subpastas). O botão
   **Testar conexão** confere as credenciais, as permissões e o acesso a até três caixas.
2. **Análises de e-mail › Nova análise** — escolha as conexões, as listas e onde procurar:
   **assunto**, **corpo**, **nomes dos anexos**, **conteúdo dos anexos** e, se quiser, remetente e
   destinatários. Opções: somente mensagens recebidas a partir de uma data, incluir a **Lixeira**
   (Itens Excluídos, marcada por padrão) e o **Lixo Eletrônico** (spam, desmarcado), tamanho máximo
   por mensagem (maiores: só o início é baixado) e downloads em paralelo.
3. **Relatório** — números da análise (caixas, mensagens, anexos lidos), gráficos de termos,
   **caixas**, **remetentes** e **onde foi encontrado** (clique numa barra para filtrar), e a lista de
   mensagens com ocorrências. Cada mensagem mostra caixa, pasta, datas de recebimento e envio,
   remetente, destinatários, anexos (com a situação da leitura de cada um), Message-ID, o link para
   abrir no Outlook na Web (Microsoft 365) e os trechos encontrados — em anexos, com o nome do anexo
   e a página, planilha ou slide. Exportações: **Excel** (abas *Resumo*, *Mensagens*, *Ocorrências*,
   *Raio-X* e *Erros*), **CSV**, **HTML** e **JSON**.
4. **Raio-X das caixas** — uma aba do relatório com um retrato de **todas as mensagens analisadas**
   (não só as que têm ocorrências), montado enquanto a análise processa as caixas: **quantidade de
   mensagens por pasta** (com a mais antiga e a mais recente de cada pasta), o **e-mail mais antigo**
   e o **mais recente** (com a caixa e a pasta), a **distribuição ao longo do tempo** (por mês; por
   ano quando o período é longo) e, quando há mais de uma caixa, a **quantidade por caixa**. Respeita
   as pastas ignoradas e, se a análise usar "recebidas a partir de", cobre esse período. Também sai
   nas exportações (aba *Raio-X* no Excel e seção no HTML e no JSON).

Nas mensagens, "quem interagiu" é o **remetente** (quem enviou a informação), os **destinatários** e
o **dono da caixa** em que a mensagem está guardada.

Todas as pastas de cada caixa são percorridas, com as subpastas. O conteúdo dos anexos é lido pelos
mesmos leitores dos arquivos (Word, Excel, PowerPoint, PDF, OpenDocument, RTF, textos, HTML e nomes
dentro de .zip), inclusive **mensagens encaminhadas como anexo** (.eml e .msg) e os anexos delas. A
análise em si é **somente leitura**: nenhuma mensagem é alterada, movida ou marcada como lida. A
exclusão das mensagens encontradas é opcional e precisa ser liberada em cada conexão (veja
*Exclusão dos itens encontrados*).

### Microsoft 365 (Exchange Online)

O acesso é pela API Microsoft Graph, com **OAuth 2.0** do Microsoft Entra ID — a Microsoft não
aceita senha. No formulário da conexão, em *Autenticação na Microsoft*, escolha uma das três formas
(detalhes em [Autenticação na Microsoft](#autenticação-na-microsoft-oauth-20)):

| Forma | Alcança | Quem autoriza |
|---|---|---|
| **Aplicativo com segredo do cliente** | Todas as caixas do locatário (ou as liberadas pelo RBAC para aplicativos). | Um administrador concede as permissões de aplicativo. |
| **Aplicativo com certificado** | O mesmo; um certificado no lugar do segredo (recomendado pela Microsoft: não é um texto que possa vazar). | O mesmo. |
| **Conta Microsoft conectada** | A caixa da conta que entrou e as caixas compartilhadas (ou de outras pessoas) às quais ela tem Acesso Total. Serve também para contas pessoais (Outlook.com, Hotmail). | A própria pessoa, ao entrar com a conta (ou o administrador, se a organização exigir). |

Para o aplicativo (segredo ou certificado), crie um registro de aplicativo no Microsoft Entra ID (é
preciso ser administrador global ou de aplicativos):

1. Em [entra.microsoft.com](https://entra.microsoft.com), abra *Identidade › Aplicativos ›
   Registros de aplicativo › Novo registro*. Nome: `CLEAN`; tipos de conta: *somente contas deste
   diretório organizacional*. Não é preciso URI de redirecionamento.
2. Em *Permissões de API › Adicionar uma permissão › Microsoft Graph › Permissões de aplicativo*,
   inclua **`Mail.Read`** e **`User.Read.All`** e clique em **Conceder consentimento do administrador**.
   (`Mail.Read` lê as mensagens; `User.Read.All` lista as caixas do locatário e encontra cada caixa
   pelo endereço de e-mail.)
3. A credencial do aplicativo:
   - **segredo**: em *Certificados e segredos › Novo segredo do cliente*, escolha a validade e copie
     o **Valor** (ele só é exibido uma vez; não confunda com o *ID do segredo*). Anote a data de
     expiração: ao vencer, gere outro e atualize a conexão no CLEAN;
   - **certificado**: no CLEAN, clique em **Gerar certificado** e em **Baixar certificado (.cer)**;
     no Entra ID, em *Certificados e segredos › Certificados › Carregar certificado*, envie o arquivo.
4. Na página *Visão geral*, copie o **ID do aplicativo (cliente)** e o **ID do diretório
   (locatário)**.
5. No CLEAN, em *Caixas de e-mail › Nova conexão › Microsoft 365*, informe os dois IDs e o segredo
   (ou o certificado), escolha *Todas as caixas do locatário* ou informe as caixas, e clique em
   *Testar conexão*.

A permissão de aplicativo `Mail.Read` dá acesso de leitura a **todas** as caixas do locatário. Para
limitar o CLEAN a algumas caixas, não conceda `Mail.Read` no Entra ID e use o **RBAC para
aplicativos** do Exchange Online (PowerShell com o módulo *ExchangeOnlineManagement*):

```powershell
Connect-ExchangeOnline
# ID do aplicativo e ID de objeto da entidade de serviço (Entra ID › Aplicativos empresariais › CLEAN)
New-ServicePrincipal -AppId <ID do aplicativo> -ObjectId <ID de objeto> -DisplayName 'CLEAN'
New-ManagementScope -Name 'Caixas analisadas pelo CLEAN' -RecipientRestrictionFilter "MemberOfGroup -eq '<DN do grupo com as caixas>'"
New-ManagementRoleAssignment -App <ID do aplicativo> -Role 'Application Mail.Read' -CustomResourceScope 'Caixas analisadas pelo CLEAN'
```

Observações: usuários sem licença do Exchange (sem caixa) e, com o RBAC para aplicativos, caixas
fora do escopo liberado são ignorados e aparecem como aviso no registro da análise; o **arquivo
morto online** (In-Place Archive) não é acessível pela API; o Exchange Online aceita até 4 downloads
simultâneos por caixa e, ao atingir o limite de requisições, o CLEAN espera o tempo indicado pelo
serviço e continua (aviso no registro). Uma pasta que não pode ser lida vai para a aba *Erros* e as
demais pastas da caixa continuam sendo analisadas.

### Autenticação na Microsoft (OAuth 2.0)

As três formas usam o protocolo OAuth 2.0 da plataforma de identidade da Microsoft; o CLEAN nunca
recebe nem guarda a senha de ninguém. Os segredos, as chaves dos certificados e as autorizações das
contas conectadas ficam gravados cifrados (veja *Credenciais e rede*) e não voltam para o navegador.

**Aplicativo com segredo do cliente.** O fluxo *client credentials*: o CLEAN apresenta o ID do
aplicativo e o segredo e recebe um token de acesso com as permissões de aplicativo concedidas pelo
administrador. Simples de configurar, mas o segredo vence (no máximo em 2 anos) e precisa ser
trocado no CLEAN.

**Aplicativo com certificado.** O mesmo fluxo, com uma asserção assinada (JWT PS256) pela chave
privada de um certificado no lugar do segredo — a forma recomendada pela Microsoft. **Gerar
certificado** cria no servidor do CLEAN uma chave RSA de 2048 bits e um certificado autoassinado
válido por 2 anos; a chave privada fica só no CLEAN (cifrada) e o arquivo `.cer` baixado tem só a
parte pública, que é o que se envia ao registro do aplicativo. Para usar um certificado da própria
empresa, escolha *Usar um certificado existente* e envie um arquivo PEM com o certificado e a chave
privada sem senha (um `.pfx` pode ser convertido com
`openssl pkcs12 -in certificado.pfx -out certificado.pem -nodes`). A tela mostra a impressão digital
(a mesma exibida no Entra ID) e a validade, e a lista de conexões avisa quando faltam menos de 30
dias para vencer. Os repositórios do OneDrive e do SharePoint ligados à conexão usam o mesmo
certificado.

**Conta Microsoft conectada.** Permissões delegadas: o CLEAN acessa o que a própria conta acessa.
Configure o registro do aplicativo assim (pode ser o mesmo do aplicativo, com as permissões
delegadas a mais):

1. Em *Registros de aplicativo › Novo registro*, escolha os tipos de conta: *somente contas deste
   diretório organizacional* (contas de trabalho ou escola) ou, para contas pessoais (Outlook.com,
   Hotmail), *contas em qualquer diretório organizacional e contas Microsoft pessoais*.
2. Em *Autenticação*, ative **Permitir fluxos de clientes públicos** (*Sim*) e salve — é o que permite
   a entrada pelo código de dispositivo. Não é preciso URI de redirecionamento.
3. Em *Permissões de API › Adicionar uma permissão › Microsoft Graph › Permissões delegadas*, inclua
   **`User.Read`**, **`Mail.Read`** e **`Mail.Read.Shared`** (para a exclusão, **`Mail.ReadWrite`** e
   **`Mail.ReadWrite.Shared`**). A própria pessoa autoriza essas permissões ao entrar; se a
   organização não permitir que usuários autorizem aplicativos, o administrador clica em *Conceder
   consentimento do administrador*.
4. No CLEAN, informe o **ID do locatário** (o GUID ou o domínio; também `organizations`, para
   qualquer conta de trabalho ou escola, ou `consumers`, para contas pessoais) e o **ID do
   aplicativo**, clique em **Conectar conta** e siga as instruções: abra
   [microsoft.com/devicelogin](https://microsoft.com/devicelogin) em qualquer navegador (inclusive no
   celular), digite o código mostrado e entre com a conta. A tela mostra a conta conectada; salve a
   conexão.

Com a conta conectada, as caixas analisadas são as da lista: a caixa da própria conta (lida em
`/me`) e caixas compartilhadas ou de outras pessoas às quais ela tenha **Acesso Total** no Exchange
Online (em branco, só a caixa da conta). Caixas compartilhadas não existem em contas pessoais, e com
`common` ou `consumers` a permissão `Mail.Read.Shared` não é pedida — para elas, informe o locatário.

A Microsoft devolve um novo token de atualização a cada uso; o CLEAN grava o novo (cifrado) a cada
análise, teste ou exclusão. A autorização vence depois de **90 dias sem uso** (ou antes, se a senha
for trocada, as sessões forem encerradas ou o acesso condicional exigir nova verificação): nesse
caso, o erro da análise pede para entrar de novo — edite a conexão e clique em *Conectar conta*.
Para a exclusão das mensagens, a entrada precisa ter autorizado a escrita: marque *Permitir
excluir* **antes** de clicar em *Conectar conta* (a tela avisa quando a conta conectada autorizou só
a leitura). Trocar o locatário, o aplicativo ou o tipo da conexão exige conectar a conta de novo.
Os repositórios do OneDrive e do SharePoint não usam a conta conectada (as permissões dela são só de
e-mail).

Mensagens comuns na entrada: *Permitir fluxos de clientes públicos* desativado (AADSTS7000218);
aplicativo de um só locatário com `organizations`/`common` (AADSTS50194: informe o ID do
locatário); permissões que precisam do administrador (AADSTS90094/65001); bloqueio por acesso
condicional (AADSTS53003). Cada uma aparece na tela com a providência a tomar.

### Google Workspace (Gmail)

1. No [Google Cloud Console](https://console.cloud.google.com), crie ou escolha um projeto e ative
   a **Gmail API** e a **Admin SDK API** (*APIs e serviços › Biblioteca*).
2. Em *IAM e administrador › Contas de serviço*, crie uma conta (ex.: `clean`). Em *Chaves ›
   Adicionar chave › Criar nova chave › JSON*, baixe o arquivo. (Se a organização bloqueia chaves de
   contas de serviço, um administrador precisa liberar a política
   `iam.disableServiceAccountKeyCreation` para o projeto.)
3. No [Admin Console](https://admin.google.com), abra *Segurança › Acesso e controle de dados ›
   Controles de API › Delegação em todo o domínio › Adicionar novo*. Informe o **ID do cliente** da
   conta de serviço (campo `client_id` do JSON) e os escopos:

   ```
   https://www.googleapis.com/auth/gmail.readonly,https://www.googleapis.com/auth/admin.directory.user.readonly
   ```

   Para usar a exclusão, acrescente `https://mail.google.com/` (exclusão definitiva) ou
   `https://www.googleapis.com/auth/gmail.modify` (mover para a Lixeira).

4. No CLEAN, em *Nova conexão › Google Workspace*, envie o arquivo JSON e informe o e-mail de um
   **administrador** (usado só para listar os usuários do domínio quando todas as caixas são
   analisadas). *Testar conexão* confere a delegação e o acesso a até três caixas.

A "pasta" de cada mensagem são os seus **marcadores** (ex.: `Caixa de entrada; Clientes/2026`); as
pastas ignoradas valem para os marcadores. Usuários sem Gmail habilitado são ignorados. Mensagens
maiores que o limite de tamanho da análise não são baixadas inteiras: o corpo e os nomes dos
anexos são verificados, mas o conteúdo dos anexos não.

### Servidores IMAP

Informe o servidor, a porta e a segurança (*SSL/TLS*, porta 993, recomendado; ou *STARTTLS*, porta
143) e as caixas. Cada caixa usa o seu **login** (por padrão, o próprio e-mail) e a sua **senha**;
caixas sem senha própria usam a **senha padrão** da conexão. *Adicionar várias* inclui uma lista de
e-mails de uma vez. Para servidores internos com certificado próprio, marque *Aceitar certificado
não confiável*.

Com uma conta de serviço, não é preciso saber a senha de cada usuário:

- **Exchange Server local** (com o serviço IMAP4 habilitado): dê à conta de serviço acesso total às
  caixas (`Add-MailboxPermission -Identity <caixa> -User svc-clean -AccessRights FullAccess
  -AutoMapping $false`) e use como login `DOMINIO\svc-clean\<alias da caixa>` com a senha da conta de
  serviço como senha padrão.
- **Dovecot** com usuário mestre: login `<caixa>*<usuário mestre>` e a senha do usuário mestre.
- **Provedores de hospedagem** e outros servidores: o login e a senha de cada caixa (ou uma senha de
  aplicativo, quando o provedor exige verificação em duas etapas).

O Gmail não aceita mais senha simples por IMAP (use o tipo Google Workspace ou uma senha de
aplicativo). No Gmail via IMAP, a pasta *Todos os e-mails* é analisada uma única vez (com os
marcadores), sem repetir as mensagens de cada marcador.

**Exchange Online e Outlook.com por IMAP (OAuth 2.0 da Microsoft).** A Microsoft desativou o login
com senha no IMAP: em *Autenticação*, escolha **OAuth 2.0 da Microsoft**. O servidor é
`outlook.office365.com` (porta 993, SSL/TLS) e o login de cada caixa é o próprio e-mail dela, com
um token da Microsoft (XOAUTH2) no lugar da senha — com as mesmas três formas do Microsoft 365:

- **conta conectada** (o mais comum): no registro do aplicativo, ative *Permitir fluxos de clientes
  públicos* e inclua a permissão delegada **`IMAP.AccessAsUser.All`** (*Microsoft Graph › Permissões
  delegadas*); clique em *Conectar conta* e entre com a conta. Analisa a caixa da conta e as caixas
  a que ela tem Acesso Total (informe-as na lista; em branco, a caixa da conta);
- **aplicativo** (segredo ou certificado): inclua a permissão de aplicativo **`IMAP.AccessAsApp`**
  (*APIs que minha organização usa › Office 365 Exchange Online*), conceda o consentimento do
  administrador e, no Exchange Online, registre o aplicativo e dê a ele Acesso Total a cada caixa:

  ```powershell
  Connect-ExchangeOnline
  New-ServicePrincipal -AppId <ID do aplicativo> -ObjectId <ID de objeto> -DisplayName 'CLEAN'
  Add-MailboxPermission -Identity financeiro@empresa.com.br -User <ID de objeto> -AccessRights FullAccess
  ```

O IMAP precisa estar habilitado nas caixas (centro de administração do Microsoft 365 › Usuários › a
pessoa › Email › Gerenciar aplicativos de email). O Exchange Online encerra a sessão IMAP quando o
token vence: o CLEAN abre cada sessão com um token válido por ao menos 30 minutos e troca a sessão
por outra quando o token está para vencer (entre as pastas e entre os lotes de mensagens); se a
sessão cair assim mesmo, a leitura continua numa nova sessão, de onde parou (sem repetir nem pular
mensagens), e só desiste da caixa depois de cinco falhas seguidas sem ler nenhuma mensagem.

Por segurança, o CLEAN só envia o token da Microsoft aos servidores IMAP da Microsoft
(`outlook.office365.com`, `outlook.office.com` e `imap-mail.outlook.com`), com o certificado do
servidor verificado — em qualquer outro endereço, o token daria acesso à caixa a quem o recebesse. Com uma senha num desses servidores, o formulário e o
erro da análise explicam como passar para o OAuth. Para o Exchange Online, o tipo **Microsoft 365**
costuma ser a melhor escolha (identificadores que sobrevivem à mudança de pasta, links para o
Outlook na Web e o tamanho de cada mensagem sem baixá-la).

### Credenciais e rede

- Senhas, segredos do cliente, chaves privadas (inclusive a dos certificados gerados pelo CLEAN) e
  os tokens de atualização das contas Microsoft conectadas são gravados **cifrados** (AES-256-GCM) no
  `db.json`. A chave fica em `data\chave-segredos.key` (criada na primeira execução) ou na variável
  `CLEAN_SECRET_KEY`. **Faça cópia da chave junto com o backup da pasta `data`**: sem ela, as senhas
  precisam ser informadas de novo.
- As credenciais nunca voltam para o navegador: nos formulários, deixar um campo de senha em branco
  mantém a senha salva. Ao trocar o servidor IMAP, as senhas salvas são descartadas (e precisam ser
  informadas de novo), para que não sejam enviadas a outro endereço.
- A entrada de uma conta Microsoft pelo código acontece no navegador de quem entra
  (`microsoft.com/devicelogin`, em qualquer computador ou celular): o servidor do CLEAN não precisa
  de acesso a essa página, só ao `login.microsoftonline.com`. Uma entrada concluída fica na memória do
  servidor por até 2 horas, e um certificado gerado por até 24 horas, até a conexão ser salva.
- O CLEAN precisa acessar `login.microsoftonline.com` e `graph.microsoft.com` (Microsoft 365; no
  IMAP com OAuth, também `outlook.office365.com` na porta 993) ou
  `oauth2.googleapis.com`, `gmail.googleapis.com` e `admin.googleapis.com` (Google) pela porta 443.
  Para o OneDrive e o SharePoint, também `*.sharepoint.com` (o conteúdo dos arquivos é baixado de
  lá; sem esse acesso, só os nomes são verificados).
  Se a rede exige **proxy**, defina, antes de iniciar o CLEAN (Node.js 22.21 ou superior), as
  variáveis de ambiente `NODE_USE_ENV_PROXY=1` e `HTTPS_PROXY=http://proxy.empresa.local:3128` — por
  exemplo, retirando o `rem` das linhas correspondentes no `iniciar.bat`.
- Se o proxy faz **inspeção de HTTPS** (o erro fala em certificado não reconhecido, como
  `UNABLE_TO_GET_ISSUER_CERT_LOCALLY`), exporte o certificado raiz da empresa em formato PEM
  (Base-64) e defina, também antes de iniciar, `NODE_EXTRA_CA_CERTS=C:\CLEAN\certificado-empresa.pem`
  (há uma linha pronta no `iniciar.bat`).

### Listagens de e-mail (contas e mensagens)

Além das análises por termos, o menu **Listagens de e-mail** faz dois inventários **somente leitura**
(não procuram nada, não baixam o conteúdo das mensagens e não excluem nada), usando as mesmas conexões
já cadastradas em *Caixas de e-mail*:

- **Contas do domínio** — lista **todas as contas registradas no domínio**, com os dados do cadastro:
  nome, endereço principal, apelidos (aliases), situação (ativa/inativa), tipo (no Microsoft 365,
  *Member*/*Guest*; no Google, usuário/administrador), se tem licença, data de criação, último acesso,
  departamento, cargo, telefone e unidade organizacional. No **Microsoft 365** exige a permissão
  `User.Read.All` (tipo Aplicativo com consentimento do administrador, ou a mesma permissão delegada
  na conta conectada); no **Google Workspace**, é consultada pela Admin SDK em nome do administrador
  informado; no **IMAP**, que não tem um catálogo de contas, traz apenas as caixas cadastradas na
  conexão, com um aviso.
- **Mensagens por caixa** — lista **todas as mensagens de todas as pastas** das caixas escolhidas, com
  os dados de cada uma: remetente, destinatários (Para e Cc), data de recebimento e de envio, assunto,
  pasta, tamanho, se tem anexos e o Message-ID. Lê só os cabeçalhos — o corpo e os anexos não são
  baixados —, então é bem mais rápida e leve que uma análise por termos. Aceita filtrar por período
  (recebidas a partir de uma data) e incluir ou não a Lixeira e o Lixo Eletrônico.

Cada listagem gera um relatório na tela com filtros e gráficos (por situação/tipo/conexão, no caso das
contas; por caixa, pasta e remetente, no caso das mensagens) e exportação para **Excel**, **CSV**,
**HTML** e **JSON**. As listagens rodam em segundo plano (como as análises), podem ser canceladas e
ficam separadas das *Análises de e-mail*, na própria seção.

## Exclusão dos itens encontrados

Além de só analisar, o CLEAN pode **excluir** os arquivos e as mensagens de e-mail em que algum
termo das listas for encontrado — sem critério adicional: basta um termo encontrado para o item ser
excluído. Há três formas de trabalhar:

| Modo | Como usar | O que acontece |
|---|---|---|
| **Somente analisar** | *Nova análise › O que fazer com os itens encontrados › Somente analisar* (padrão) | Gera o relatório; nada é alterado. |
| **Analisar e excluir automaticamente** | *Nova análise › Analisar e excluir automaticamente* e digitar **EXCLUIR** para confirmar | Cada item encontrado é excluído durante a análise, sem confirmação item a item. Arquivos: depois de registrados no relatório com o último usuário (ao fim de cada repositório ou a cada 200 arquivos). E-mails: ao fim de cada caixa (para não atrapalhar a leitura das pastas). |
| **Excluir depois, item a item** | No relatório, abra o item e clique em **Excluir arquivo** / **Excluir mensagem** | Exclui só aquele item, depois de uma confirmação. Se o arquivo mudou depois da análise, o CLEAN avisa e pede uma segunda confirmação. |
| **Excluir depois, em lote** (arquivos) | No relatório, marque os arquivos e clique em **Excluir selecionados**, ou use **Excluir todos os filtrados** | Exclui os escolhidos em segundo plano, depois de uma prévia e da confirmação **EXCLUIR** — veja [Exclusão em lote pelo relatório](#exclusão-em-lote-pelo-relatório). |

Salvaguardas:

- a exclusão só funciona nos repositórios e nas conexões de e-mail com **Permitir exclusão**
  marcado no cadastro (desmarcado por padrão). A permissão é conferida de novo quando a análise
  começa (ela pode ter esperado na fila) e, se for desligada durante a análise, nada mais é excluído
  daquele repositório ou conexão;
- o modo automático exige digitar **EXCLUIR** ao iniciar a análise. Ao **cancelar**, as exclusões
  que já estavam em andamento terminam (e ficam registradas) e nenhuma outra começa;
- **arquivos**: só são excluídos se continuarem iguais ao que foi analisado (mesmo tamanho e data de
  modificação) — no modo automático, os alterados ficam e aparecem como *alterados depois da
  análise*; e só dentro da pasta do repositório, conferida pelo caminho real: uma pasta trocada
  depois da análise por um atalho (link simbólico ou junção) que aponte para fora do repositório
  impede a exclusão;
- os arquivos da pasta de dados e da pasta de instalação do CLEAN nunca são excluídos (a pasta de
  dados nem é analisada); um repositório cadastrado **dentro** de outro, sem *Permitir exclusão*,
  protege os seus arquivos também quando a análise é feita pelo repositório maior — no OneDrive e no
  SharePoint, um repositório sem exclusão que lista contas ou sites protege essas contas (conferidas
  no Microsoft 365, pelo e-mail ou pelo nome de logon; se a conferência falhar, nada do OneDrive é
  excluído) e esses sites, com os subsites;
- atenção ao verificar o nome com **Caminho completo**: um termo no nome de uma pasta faz todos os
  arquivos dela (e das subpastas) serem encontrados — e, no modo automático, excluídos;
- **e-mails, OneDrive e SharePoint**: vale a forma de exclusão (definitiva ou para a lixeira) que
  estava no cadastro ao criar a análise — se ela mudar para *lixeira* antes de a análise começar, a
  lixeira é usada. No relatório, se a forma mudou depois de a página ser aberta, o CLEAN avisa e pede
  nova confirmação; se a conta ou o site do arquivo saiu do cadastro do repositório (ou o tipo ou o
  locatário mudaram), ou se o arquivo está numa pasta que o repositório passou a ignorar, a exclusão
  pelo relatório deixa de estar disponível;
- cada tentativa de exclusão fica registrada: no relatório (situação de cada item, filtro
  *Exclusão* e o bloco *Excluídos*), no **Registro** da análise (exclusões manuais e falhas), na aba
  **Exclusões** do Excel (quando, o quê, automática ou manual, por quem e o resultado), no CSV
  (coluna *Exclusão*) e no HTML (situação abaixo de cada item). Há também um registro geral,
  **`data\exclusoes.ndjson`** (uma linha por exclusão, com a análise, o item, quando, como, por quem
  e o resultado), que continua existindo mesmo que o relatório seja excluído;
- "quem excluiu" é o usuário da autenticação (`AUTH_USER` — uma conta única, compartilhada por quem
  usa o CLEAN) com o endereço de acesso; atrás de um proxy na mesma máquina (IIS com ARR, por
  exemplo), o endereço do navegador informado pelo proxy. Nas exclusões automáticas, fica registrado
  quem iniciou a análise.

Situações de exclusão de cada item: **Excluído**; **Não encontrado** (o item já não existia na hora
da exclusão — excluído ou movido por outra pessoa); **Não excluído: alterado depois da análise**
(arquivos, no modo automático); **Falha na exclusão** (com o motivo). O bloco *Excluídos* conta só
os excluídos de fato; os demais aparecem ao lado.

Como cada tipo é excluído e a permissão necessária:

| Onde | Exclusão | Permissão |
|---|---|---|
| Arquivos (pastas e compartilhamentos) | **Definitiva**: arquivos apagados pela rede não vão para a Lixeira do Windows. Arquivos somente leitura também são excluídos. | A conta do CLEAN precisa de permissão de **modificação** (NTFS e compartilhamento), não só de leitura. |
| Microsoft 365 | *Excluir definitivamente* (a mensagem vai para a área de expurgo e some para o usuário) ou *Mover para a Lixeira* (Itens Excluídos), conforme a conexão. | Aplicativo: **`Mail.ReadWrite`** (tipo Aplicativo) no lugar de `Mail.Read`; com o RBAC para aplicativos, a função `Application Mail.ReadWrite`. Conta conectada: **`Mail.ReadWrite`** e **`Mail.ReadWrite.Shared`** (delegadas), autorizadas ao conectar a conta com *Permitir excluir* marcado. |
| Google Workspace | Definitiva ou para a Lixeira, conforme a conexão. | Na delegação em todo o domínio, inclua o escopo `https://mail.google.com/` (definitiva) ou `https://www.googleapis.com/auth/gmail.modify` (lixeira). |
| OneDrive e SharePoint | *Mover para a Lixeira* do site ou do OneDrive (padrão; o usuário ou o administrador do site pode restaurar por até 93 dias) ou *Excluir definitivamente*, conforme o repositório. Só se o arquivo continuar como foi analisado (mesma versão, mesmo nome e mesma pasta — conferido antes e, pelo cabeçalho If-Match, também na própria exclusão); arquivos abertos para edição, em check-out ou com rótulo de retenção (registro) não são excluídos. | **`Files.ReadWrite.All`** (ou `Sites.ReadWrite.All`; com `Sites.Selected`, permissão `Write` no site). |
| IMAP | *Definitiva*: marca e expurga só as mensagens encontradas (UID EXPUNGE); o servidor precisa oferecer a extensão **UIDPLUS** — sem ela, a exclusão definitiva é recusada, porque um expurgo comum apagaria também as outras mensagens marcadas como excluídas na pasta. *Para a Lixeira*: move para a pasta Lixeira do servidor (MOVE; sem MOVE, copia, confere a cópia e só então expurga, o que também exige UIDPLUS); mensagens que já estão na Lixeira ficam lá. No **Gmail via IMAP**, a exclusão definitiva move para a Lixeira e expurga de lá (expurgar de outra pasta só tiraria o marcador). | A conta precisa poder alterar a caixa. |

Importante:

- **não há como desfazer** a exclusão definitiva pelo CLEAN. Confira as listas de referência e rode
  primeiro *Somente analisar* para ver o que seria excluído;
- em e-mails, a mensagem inteira é excluída mesmo quando o termo está só em um anexo;
- antes de excluir uma mensagem por IMAP, o CLEAN confere se ela ainda é a mesma da análise
  (Message-ID e UIDVALIDITY da pasta) e, depois, se ela de fato saiu da pasta; se o servidor recusar
  a marcação ou o expurgo, a mensagem é dada como não excluída (e a marcação feita pelo CLEAN é
  desfeita);
- retenções, bloqueios de litígio (Microsoft Purview, Google Vault) e backups do provedor ou do
  servidor de arquivos (cópias de sombra, snapshots) continuam valendo: o item pode continuar
  preservado neles, como exige a política da empresa;
- o Microsoft 365 é acessado com identificadores imutáveis: uma mensagem movida de pasta depois da
  análise ainda pode ser excluída pelo relatório. Se a resposta de uma exclusão se perder (falha de
  rede) e a nova tentativa não encontrar a mensagem, ela é dada como excluída.

### Exclusão em lote pelo relatório

Nos relatórios de arquivos (análises por termos, buscas por tipo e execuções das políticas de
retenção), depois de revisar os resultados:

1. marque os arquivos (a caixa ao lado de cada um; *Selecionar todos* marca os da página) e clique
   em **Excluir selecionados** — a seleção vale entre as páginas e é desfeita quando os filtros
   mudam —, ou use **Excluir todos os filtrados** (sem filtros, **Excluir todos os arquivos**);
2. a confirmação mostra quantos arquivos serão excluídos em cada repositório e de que forma
   (definitiva nas pastas do Windows; para a lixeira ou definitiva no OneDrive e no SharePoint,
   conforme o repositório) e quantos ficam de fora: já excluídos, em repositórios sem *Permitir
   exclusão*, em locais protegidos, de repositórios removidos ou alterados no cadastro, em pastas
   que passaram a ser ignoradas. Digite **EXCLUIR**;
3. a exclusão roda em segundo plano, um arquivo de cada vez, e o relatório mostra o andamento — dá
   para fechar a página e voltar depois. **Parar a exclusão** interrompe o lote depois do arquivo em
   andamento.

Cada arquivo passa pelas mesmas conferências da exclusão item a item, mas sem a segunda
confirmação: os alterados depois da análise (na retenção, também os que deixaram de estar
expirados) **são mantidos** e aparecem como alterados. A confirmação vale para os arquivos que a
prévia mostrou: se os filtros ou a seleção mudarem enquanto ela é calculada, ou se mais arquivos
passarem a poder ser excluídos até o EXCLUIR (uma pasta que deixou de ser ignorada, por exemplo), o
CLEAN pede para começar de novo; se a forma de exclusão de um repositório mudar, também — e, se mudar
durante o lote, os arquivos daquele repositório falham em vez de serem excluídos de outra forma. A
prévia avisa quantos arquivos foram encontrados só pelo tipo real (na busca por tipo), para
conferir antes.

Os arquivos são excluídos em ordem de caminho. Arquivos do OneDrive e do SharePoint em contas ou
sites protegidos por um repositório sem *Permitir exclusão* são pulados. Depois de 5 falhas seguidas
numa pasta (uma pasta sem permissão, por exemplo), os demais arquivos dela não são tentados naquele
lote; depois de 20 falhas seguidas em várias pastas (a conta do CLEAN sem permissão de modificação,
por exemplo), o lote para, e o motivo aparece no relatório. Ao repetir o lote, os que falharam antes
ficam por último, e assim ele avança nos demais.

Cada exclusão é registrada na hora, como manual, com quem confirmou e a indicação *(exclusão em
lote)* — se a gravação do registro falhar (disco cheio, por exemplo), o lote para: uma exclusão não
fica sem rastro. O **Registro** da análise mostra o início e o resultado de cada lote. Um lote por
relatório de cada vez (um segundo pedido, de outra aba, é recusado), com até 100.000 arquivos
(filtre o relatório para excluir mais); enquanto ele roda, o relatório não pode ser excluído. Ao
encerrar, o CLEAN não começa novos lotes e espera o arquivo em exclusão terminar; o lote para ali (o
que já foi excluído fica registrado, e o Registro da análise anota a interrupção) e pode ser
iniciado de novo com os arquivos que restaram. Nos relatórios de e-mail, a exclusão continua item a
item.

## Agendamentos

Em **Automação › Agendamentos**, cadastre análises de arquivos ou de e-mail que o CLEAN executa
sozinho nos dias e horários definidos. O formulário é o mesmo da nova análise (locais, listas — ou
tipos de arquivo, na [busca por tipo](#busca-por-tipo-de-arquivo) —, o que verificar e o que fazer
com os itens encontrados), com a regra de recorrência; na tela *Nova análise*, a opção **Quando
executar › Agendar** faz o mesmo.

| Repetição | Exemplos |
|---|---|
| **Uma vez** | em 30/09/2026 às 22:00 |
| **A cada algumas horas** | a cada 2 horas, das 08:00 às 18:00, de segunda a sexta |
| **Diariamente** | todos os dias às 02:00; a cada 3 dias; só em dias úteis (de segunda a sexta) |
| **Semanalmente** | às segundas e quintas às 22:00; a cada 2 semanas, aos sábados |
| **Mensalmente** | no dia 1; no dia 31 (nos meses mais curtos, no último dia); no último dia do mês; na primeira segunda-feira; no último sábado; a cada 3 meses |

As repetições têm data de início e término: nunca, numa data ou **depois de um número de
execuções** — contam as execuções de fato iniciadas pelo agendamento a partir de quando ele é salvo
(horários pulados ou perdidos e *Executar agora* não contam; mudar a regra recomeça a contagem). Ao
montar a regra, a tela mostra a descrição, as **próximas cinco execuções** e a última prevista.

Como funcionam:

- os horários são os do **servidor** do CLEAN (o fuso aparece na tela) e o CLEAN precisa estar em
  execução nesses horários: instale-o como serviço (veja [Executando como serviço](#executando-como-serviço));
- cada execução entra na fila como qualquer análise (`MAX_CONCURRENT_SCANS`) e gera um relatório
  comum, com o nome do agendamento e a data; a lista de análises marca as *agendadas*;
- se uma execução do mesmo agendamento ainda estiver em andamento (ou na fila), a nova é **pulada**
  e fica registrada no histórico (duas execuções do mesmo agendamento nunca rodam juntas);
- **horário perdido** (CLEAN parado ou computador desligado): na volta, o agendamento é executado
  **uma vez** (opção marcada por padrão) ou o horário fica só registrado como perdido — nunca uma
  execução para cada horário perdido;
- **Executar agora** roda o agendamento na hora, sem mudar a próxima execução programada;
  **Pausar** suspende o agendamento (os horários da pausa não são executados depois); se o relógio
  do servidor for corrigido para trás, a próxima execução é recalculada;
- o **Histórico** mostra as últimas 50 execuções: quando, se foi no horário, atrasada ou manual, o
  resultado (concluída, pulada, não iniciada e o motivo, perdida), quantos itens foram encontrados e
  excluídos e o link do relatório;
- **relatórios guardados**: é possível manter só os mais recentes de cada agendamento. Quando uma
  execução termina, ficam os relatórios até o N-ésimo **concluído** mais recente (os que falharam
  não contam) e sempre o da última análise completa; os mais antigos são excluídos (o registro geral
  `data\exclusoes.ndjson` é mantido);
- repositórios, conexões de e-mail e listas usados por um agendamento não podem ser excluídos do
  cadastro enquanto fizerem parte dele.

**Itens analisados em cada execução:**

- **todos** os arquivos (ou mensagens);
- os alterados (ou recebidas) nos **últimos N dias**;
- **incremental**: a partir da segunda execução, só os arquivos modificados, criados ou copiados para
  o repositório — ou as mensagens recebidas — desde o início da última execução **concluída sem
  falhas de acesso** (uma execução que não conseguiu ler um repositório, uma conta, um site, uma
  conexão ou uma caixa inteira não serve de base), com 1 hora de margem para diferenças de relógio
  entre os servidores. A primeira execução é completa, assim como a seguinte a qualquer mudança nos
  locais, nas listas de referência (um termo novo precisa ser procurado em tudo), nas opções ou na
  ação (ao passar a excluir). A **análise completa periódica** (a cada 2 a 50 execuções; 7 por
  padrão) é obrigatória: ela pega o que as incrementais não veem — pastas movidas inteiras para o
  repositório (os arquivos mantêm as datas), itens movidos no OneDrive e no SharePoint, mensagens
  movidas entre pastas ou importadas (de um .pst, por exemplo), itens com erro de leitura ou de
  exclusão. Cada relatório incremental mostra só o que foi encontrado no seu período.

**Exclusão automática agendada:**

- exige, a cada vez que o agendamento é salvo, **Permitir exclusão** em todos os locais e a
  confirmação digitada (**EXCLUIR**); o CLEAN registra quem confirmou, quando, o alcance e a forma
  de exclusão de cada local e os **critérios** naquele momento: os termos das listas de referência,
  as pastas, contas, sites, caixas e pastas de e-mail ignorados e os locais protegidos por
  repositórios sem *Permitir exclusão*;
- em cada execução, a confirmação é conferida com o cadastro atual: se um repositório ou conexão
  deixou de permitir a exclusão, mudou de caminho, de contas ou de sites (OneDrive e SharePoint), de
  conta, servidor ou caixas (e-mail), passou a excluir de forma definitiva (antes, para a lixeira),
  ou se os critérios mudaram (um termo novo numa lista, uma pasta que deixou de ser ignorada, um
  repositório protegido removido), a execução **não é iniciada** e o motivo aparece no histórico e
  na lista de agendamentos até o agendamento ser salvo e confirmado de novo. Ao salvar uma lista,
  um repositório ou uma conexão com uma alteração dessas, o CLEAN avisa quais agendamentos ficaram
  suspensos;
- a conferência também é feita quando uma execução que esperou na fila começa: se nesse intervalo
  algo mudou (inclusive as caixas de uma conexão de e-mail) ou o agendamento foi pausado, excluído ou
  passou a *Somente analisar*, ela analisa sem excluir. Pausar, excluir ou mudar o agendamento também
  interrompe a exclusão de uma execução dele em andamento;
- nas exclusões, "quem excluiu" fica registrado como o agendamento e quem confirmou (ex.:
  `agendamento "Limpeza semanal" (exclusão automática confirmada por acesso local em 25/09/2026 10:00)`);
- *Executar agora* num agendamento com exclusão pede uma confirmação.

## Políticas de retenção

Em **Automação › Retenção**, crie políticas que eliminam os arquivos e as mensagens de e-mail **mais
antigos que uma idade máxima** — por exemplo, *arquivos sem uso há mais de 5 anos* ou *e-mails
recebidos há mais de 10 anos*. Aqui não há listas de referência: vale só a data do item.

**Critério de data dos arquivos:**

| Critério | Data considerada | Quando usar |
|---|---|---|
| **Sem uso** (padrão) | A mais recente entre a última modificação, o último acesso e a criação | O mais seguro: o arquivo só expira se **nenhuma** dessas datas for recente. Sem o registro do último acesso (veja abaixo), um arquivo só lido não tem data recente — ele fica protegido só se tiver sido modificado ou criado recentemente. |
| **Última modificação** | Quando o conteúdo foi alterado pela última vez | Confiável em qualquer servidor. Não considera os arquivos que só foram abertos (lidos). |
| **Último acesso (abertura)** | Quando o arquivo foi aberto pela última vez | Só se o servidor de arquivos registrar o último acesso (veja abaixo). Não existe no OneDrive e no SharePoint. |
| **Criação** | Quando o arquivo foi criado **ou copiado** para o repositório | Pastas de passagem: digitalizações, exportações, arquivos temporários. |

Nas mensagens de e-mail, vale a **data de recebimento** (nos Itens Enviados, a data do envio).

**Último acesso no Windows.** O NTFS só atualiza a data do último acesso se esse registro estiver
ligado — em muitos servidores ele fica desligado para economizar gravações em disco. Para conferir,
rode no servidor de arquivos, num prompt de comando como administrador:

```bat
fsutil behavior query disablelastaccess
```

`0` ou `2`: o último acesso é registrado; `1` ou `3`: não é (para ligar,
`fsutil behavior set disablelastaccess 0`; a mudança pode exigir reiniciar o servidor). Sem o
registro, a data fica parada — em geral na criação ou na cópia do arquivo — e um arquivo aberto
todos os dias parece antigo: pelo critério *Último acesso*, ele seria excluído. No sentido oposto,
programas que leem todos os arquivos (antivírus, backup, indexação e as **análises de conteúdo do
próprio CLEAN**) podem atualizar o último acesso de tudo, e aí nada expira pelos critérios *Último
acesso* e *Sem uso*; nos repositórios analisados com frequência, prefira *Última modificação*. Depois
de uma migração de servidor, a data de criação costuma ser a da cópia: com *Sem uso* ou *Criação*,
nada expira até essa data ficar antiga.

**Idade máxima:** em dias, meses ou anos (meses e anos pelo calendário). A **data de corte** é
calculada no início de cada execução — expiram os itens com a data do critério **anterior** a ela —
e fica registrada no relatório; o formulário mostra a data de corte se a política fosse executada
hoje. Itens sem a data do critério (um sistema de arquivos sem data de criação, por exemplo) — ou
com datas anteriores a 02/01/1980, que são valores padrão de sistemas antigos e de datas perdidas
(01/01/1970, 01/01/1980) — nunca expiram e são contados à parte.

**Demais opções:**

- arquivos: **somente os nomes** que combinam com padrões (`*.tmp`, `*.bak`, `~$*` — um por linha, só
  o nome, sem pastas); em branco, todos os arquivos. Pastas e arquivos ignorados no cadastro do
  repositório continuam ignorados;
- e-mail: incluir ou não a **Lixeira** (Itens Excluídos) e o **Lixo Eletrônico**; as pastas
  ignoradas no cadastro da conexão continuam ignoradas. Com a exclusão *para a lixeira*, as
  mensagens expiradas que já estão na Lixeira aparecem no relatório (marcadas *na lixeira*), mas não
  são movidas de novo nem contam como excluídas: o provedor as apaga pela regra da própria Lixeira;
- **o que fazer**: *Somente listar (simulação)* ou *Excluir os itens expirados*. A exclusão exige
  **Permitir exclusão** nos locais escolhidos e a confirmação **EXCLUIR** a cada vez que a política é
  salva, e é conferida em cada execução como nos [agendamentos com exclusão](#agendamentos): se um
  local deixar de permitir a exclusão ou mudar de caminho, contas, sites ou caixas, a execução não é
  iniciada até a política ser salva e confirmada de novo; uma execução que esperou na fila não exclui
  se, nesse intervalo, a política foi alterada, pausada ou excluída;
- **forma de exclusão**: *Definitiva* (padrão) ou *Para a lixeira*. Vale para o e-mail, o OneDrive e o
  SharePoint (a forma da política, e não a do cadastro); nas pastas do Windows a exclusão é sempre
  definitiva;
- **limite de exclusões por execução** (1.000 por padrão; 0 = sem limite): um freio contra uma
  regra errada. Contam os itens de fato excluídos (uma falha não gasta a vaga) e as falhas têm o
  mesmo limite, para a mesma falha não se repetir milhares de vezes (permissão, arquivo em uso,
  rótulo de retenção). Ao atingir um dos dois, a execução para de excluir e avisa; os demais itens
  expirados ficam só no relatório e são excluídos nas execuções seguintes. As exclusões seguem a
  ordem da varredura (não necessariamente dos mais antigos para os mais novos). No e-mail, as
  exclusões de cada caixa são feitas ao fim dela: as vagas de mensagens que falharam valem para as
  caixas seguintes (não para o restante da mesma caixa). Itens em locais
  protegidos (um repositório sem *Permitir exclusão* dentro do analisado, as pastas do CLEAN, contas
  ou sites protegidos) aparecem no relatório, mas nunca são tentados nem contam no limite. Se não
  for possível conferir as contas protegidas do OneDrive (falha no Microsoft 365), nada daquele
  repositório é excluído e cada item conta como falha;
- **quando executar**: *Manualmente* (pelos botões da lista) ou *Agendar*, com as mesmas regras de
  recorrência, horários perdidos, histórico e relatórios guardados dos agendamentos (os relatórios
  das simulações contam à parte: simular não apaga o relatório de uma execução que excluiu).

**Na lista de políticas:** **Simular agora** executa a política sem excluir nada — o relatório mostra
exatamente o que seria excluído (faça isso antes de ligar a exclusão); **Executar e excluir agora**
(nas políticas com exclusão) pede uma confirmação; também há pausar e retomar (agendadas), histórico,
editar e excluir. As políticas agendadas aparecem no painel, em *Próximas execuções agendadas*.

**Relatório de cada execução:** os itens expirados com a data considerada e a idade, os mais antigos
primeiro; gráficos por faixa de idade (clique para filtrar), espaço por extensão e por repositório,
últimos usuários, caixas, pastas e remetentes; a regra, a data de corte e a forma de exclusão no
cabeçalho. O Excel (abas *Resumo*, *Arquivos expirados* ou *Mensagens expiradas*, *Exclusões* e
*Erros*), o CSV e o HTML têm uma linha por item, com a data considerada, a idade e a faixa de idade.
As execuções também aparecem em *Análises de arquivos* e *Análises de e-mail*, marcadas como
*retenção*.

**Como os itens são lidos e excluídos:**

- pastas do Windows: só as datas de cada arquivo (o conteúdo não é lido); o arquivo expirado é
  excluído logo depois de registrado no relatório (com o proprietário, se a opção estiver marcada),
  com as mesmas conferências da exclusão automática — tamanho e data de modificação iguais aos da
  listagem, dentro do repositório e fora dos locais protegidos — e **se continuar expirado**: a data
  do critério é lida de novo na hora de excluir (um arquivo aberto depois da listagem é mantido,
  como *alterado depois da análise*). A exclusão manual pelo relatório faz a mesma conferência e,
  se o arquivo deixou de estar expirado, pede uma segunda confirmação;
- OneDrive e SharePoint: as datas vêm do Microsoft 365, sem baixar os arquivos; a exclusão confere a
  versão do arquivo;
- e-mail: só os cabeçalhos das mensagens anteriores à data de corte são lidos (no Microsoft 365 e no
  Gmail, o próprio servidor filtra pela data); o corpo e os anexos não são baixados;
- as exclusões ficam registradas no relatório, na aba *Exclusões* do Excel e em
  `data\exclusoes.ndjson` como **Política de retenção**, com o nome da política e quem confirmou a
  exclusão (ex.: `política de retenção "Temporários" (exclusão confirmada por acesso local em
  26/09/2026 10:00)`).

Antes de ligar a exclusão, confirme com as áreas responsáveis os **prazos legais de guarda** dos
documentos (fiscais, trabalhistas, contábeis). Retenções e bloqueios de litígio do Microsoft Purview
e do Google Vault continuam valendo sobre o que o CLEAN excluir.

## Busca por tipo de arquivo

Em **Nova análise › O que procurar › Tipos de arquivo** — ou pelo botão **Buscar por tipo de
arquivo** em *Análises de arquivos* e no painel —, o CLEAN lista os arquivos dos tipos escolhidos:
vídeos e músicas nos compartilhamentos, instaladores esquecidos, backups e temporários antigos,
imagens de disco. Não há lista de referência: vale o tipo do arquivo.

**O que procurar** (pode combinar vários):

| Tipo | Extensões (exemplos; a lista completa aparece na tela) |
|---|---|
| Vídeos | .mp4, .mov, .avi, .mkv, .wmv, .mpg, .webm, .3gp, .vob, .m2ts |
| Músicas e áudio | .mp3, .wav, .wma, .aac, .m4a, .flac, .ogg, .opus, .mid |
| Imagens e fotos | .jpg, .png, .gif, .bmp, .tif, .webp, .heic, .psd e fotos RAW (.cr2, .nef, .arw, .dng) |
| Executáveis e instaladores | .exe, .msi, .msix, .bat, .cmd, .ps1, .vbs, .scr, .jar, .apk, .dmg |
| Compactados | .zip, .rar, .7z, .tar, .gz, .tgz, .bz2, .xz, .cab |
| Imagens de disco e máquinas virtuais | .iso, .img, .vhd, .vhdx, .vmdk, .vdi, .qcow2, .ova, .wim |
| Temporários e backups | .tmp, .temp, .bak, .old, .orig, .dmp, .crdownload, .part |
| *Arquivos de trabalho:* documentos, planilhas, apresentações, e-mails e bancos de dados | .doc, .docx, .pdf, .txt, .xls, .xlsx, .csv, .ppt, .pptx, .pst, .ost, .msg, .eml, .mdb, .accdb, .sqlite, .mdf |

- **Outras extensões**: as que você informar, separadas por vírgula ou espaço (`.dwg, .log`),
  sozinhas ou junto com os tipos. Vale o fim do nome, sem diferenciar maiúsculas, inclusive
  extensões compostas como `.tar.gz`;
- **Tamanho mínimo** (MB): só os arquivos a partir desse tamanho — por exemplo, os vídeos de 100 MB
  ou mais. Os menores do tipo procurado são contados à parte no relatório;
- **Conferir o tipo real pelo conteúdo** (pastas do Windows, para os tipos marcados): acha arquivos
  **renomeados** — um vídeo salvo como `relatorio.pdf`, uma foto sem extensão — lendo os primeiros
  4 KB dos arquivos **sem extensão** ou com a extensão de **outro tipo da lista**. Reconhece vídeos,
  áudios, imagens (inclusive HEIC e as fotos RAW da Canon, .CR3), executáveis, compactados, discos
  virtuais (VHDX, VHD, VMDK, QCOW e VDI), PDF, arquivos de dados do Outlook e bancos de dados SQLite
  e Access; planilhas, apresentações, textos, documentos antigos do Office, imagens ISO e
  temporários são achados só pela extensão. Para evitar falsos positivos: arquivos com extensões
  **fora da lista** valem pelo que são (um `.dll` é um executável, um `.ai` do Illustrator é um PDF
  por dentro, um `.lrcat` do Lightroom é um banco SQLite); o formato próprio de uma extensão não conta
  como renomeado (um `.docx` ou um `.doc` com conteúdo ZIP, um `.m4a`, um `.heic` ou um `.cr3` em
  MPEG-4, um `.ova` em TAR — e, em geral, um contêiner de áudio e vídeo com a extensão de outro
  áudio, vídeo ou imagem); as assinaturas curtas são conferidas com o que vem depois (o cabeçalho de
  um executável, dois quadros seguidos de MP3, vários pacotes de vídeo MPEG-TS) e textos (inclusive
  em UTF-16) não são confundidos com elas. Mesmo assim, o tipo real é uma pista: **os arquivos
  encontrados só por ele nunca são excluídos automaticamente** — ficam no relatório (filtro
  *Encontrado por: tipo real*) para a revisão. A conferência deixa a busca mais lenta (cada arquivo é
  aberto) e, como a leitura do conteúdo nas análises por termos, pode atualizar o último acesso dos
  arquivos. No OneDrive e no SharePoint vale só a extensão (o conteúdo não é baixado);
- somente os arquivos alterados a partir de uma data, os repositórios, os arquivos processados em
  paralelo e o proprietário (NTFS), como numa análise por termos.

**O que fazer com os arquivos encontrados:**

- **Somente procurar** (padrão): gera o relatório para a revisão. Depois, exclua pelo relatório —
  item a item, os selecionados ou todos os filtrados (veja
  [Exclusão em lote pelo relatório](#exclusão-em-lote-pelo-relatório));
- **Procurar e excluir automaticamente**: cada arquivo encontrado é excluído durante a busca, sem
  confirmação item a item, com as salvaguardas da [exclusão automática](#exclusão-dos-itens-encontrados)
  (*Permitir exclusão*, **EXCLUIR**, arquivos alterados depois de listados são mantidos, locais
  protegidos nunca são tentados) e um **limite de exclusões por execução** (1.000 por padrão; 0 = sem
  limite), que funciona como o das [políticas de retenção](#políticas-de-retenção): acima dele, os
  arquivos encontrados ficam só no relatório e o CLEAN avisa. Os encontrados só pelo tipo real ficam
  para a revisão (não são excluídos automaticamente). A tela avisa quando arquivos de trabalho (os
  tipos ou as extensões digitadas) estão marcados junto com a exclusão automática.

Rode primeiro *Somente procurar* e revise o relatório antes de ligar a exclusão automática — em
especial com os arquivos de trabalho e com *Conferir o tipo real*.

**Agendada:** em **Automação › Agendamentos › Agendar busca por tipo** (ou *Quando executar ›
Agendar*), com a mesma recorrência, período (todos os arquivos, os alterados nos últimos dias ou
incremental), histórico e relatórios guardados dos [agendamentos](#agendamentos). Na exclusão
automática agendada, os tipos, as extensões de cada tipo, o tamanho mínimo, o limite e a versão do
reconhecimento pelo conteúdo ficam registrados na confirmação: se forem alterados (inclusive por uma
atualização do CLEAN que mude o que um tipo abrange), é preciso salvar e confirmar de novo (uma
execução que esperou na fila procura sem excluir). Na incremental, uma execução que parou no limite
de exclusões não serve de base: a seguinte volta a procurar em tudo, para excluir o que ficou.

**Relatório:** arquivos verificados e encontrados, o espaço ocupado pelos encontrados, quantos foram
achados pelo tipo real e o espaço já excluído; gráficos por tipo, espaço por extensão, últimos
usuários e espaço por repositório (clique para filtrar); filtros por tipo, *encontrado por* (extensão
ou conteúdo), usuário, repositório, extensão e exclusão, com os maiores arquivos primeiro. O Excel
(abas *Resumo*, *Arquivos encontrados*, *Exclusões* e *Erros*), o CSV e o HTML têm uma linha por
arquivo, com o tipo, a extensão, como ele foi encontrado e o formato real. As buscas aparecem em
*Análises de arquivos* marcadas como *busca por tipo*.

## Executando como serviço

Para que o CLEAN inicie com o Windows, sem sessão aberta — necessário para os
[agendamentos](#agendamentos) e as [políticas de retenção](#políticas-de-retenção) agendadas —, use
o Agendador de Tarefas (nativo). Em um PowerShell como administrador:

```powershell
$acao = New-ScheduledTaskAction -Execute 'C:\Program Files\nodejs\node.exe' -Argument 'src\server.js' -WorkingDirectory 'C:\CLEAN'
$gatilho = New-ScheduledTaskTrigger -AtStartup
$config = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName 'CLEAN' -Action $acao -Trigger $gatilho -Settings $config -User 'EMPRESA\svc-clean' -Password 'senha-da-conta'
Start-ScheduledTask -TaskName 'CLEAN'
```

Alternativas: [NSSM](https://nssm.cc) ou [WinSW](https://github.com/winsw/winsw), que registram o
`node.exe src\server.js` como serviço do Windows com a conta escolhida.

## Configuração

As opções são variáveis de ambiente, que também podem ficar em um arquivo `.env` na pasta do
projeto (copie o `.env.example`):

| Variável | Padrão | Descrição |
|---|---|---|
| `PORT` | `3000` | Porta HTTP. |
| `HOST` | `127.0.0.1` | Endereço de escuta. Use `0.0.0.0` para permitir acesso por outros computadores. |
| `ALLOWED_HOSTS` | — | Outros nomes de acesso aceitos, separados por vírgula (apelido DNS, endereço do proxy). `localhost`, o nome e os IPs da máquina já são aceitos. |
| `AUTH_USER` / `AUTH_PASSWORD` | — | Ativa usuário e senha (autenticação HTTP básica) para a interface e a API. |
| `DATA_DIR` | `data` | Pasta com a configuração (`db.json`) e os resultados de cada análise (`scans\<id>`). |
| `MAX_CONCURRENT_SCANS` | `1` | Análises simultâneas; as demais aguardam na fila. |
| `POWERSHELL_PATH` | `powershell.exe` | PowerShell usado para o proprietário e o log de auditoria. |
| `CLEAN_SECRET_KEY` | — | Chave (32 bytes em base64) que cifra as senhas das caixas de e-mail. Sem ela, é usada a chave do arquivo `data\chave-segredos.key`. |
| `NODE_USE_ENV_PROXY` / `HTTPS_PROXY` | — | Proxy de saída para o Microsoft 365 e o Google (precisam estar definidas antes de iniciar o Node.js; veja *Credenciais e rede*). |

## Segurança

Os relatórios mostram nomes de arquivos, usuários e **trechos do conteúdo** — incluindo, conforme a
lista, dados pessoais e senhas encontradas. Por isso:

- por padrão o servidor só aceita conexões do próprio computador (`HOST=127.0.0.1`);
- só são atendidos pedidos endereçados a nomes conhecidos (localhost, o nome e os IPs da máquina e
  os de `ALLOWED_HOSTS`), o que bloqueia ataques de *DNS rebinding* vindos de outros sites;
- ao liberar o acesso pela rede, defina `AUTH_USER` e `AUTH_PASSWORD` e, de preferência, publique o
  CLEAN atrás de um proxy HTTPS (IIS com URL Rewrite/ARR, por exemplo) e inclua o endereço público
  em `ALLOWED_HOSTS`;
- proteja a pasta `data` com permissões NTFS restritas (ela guarda os resultados das análises, as
  credenciais cifradas das caixas de e-mail e dos repositórios do OneDrive/SharePoint e a chave que
  as decifra), por exemplo:
  `icacls C:\CLEAN\data /inheritance:r /grant:r "Administradores:(OI)(CI)F" "EMPRESA\svc-clean:(OI)(CI)M"`;
- dê à conexão de e-mail apenas o acesso necessário (permissões de leitura; no Microsoft 365, de
  preferência limitado às caixas analisadas pelo RBAC para aplicativos, ou uma conta conectada que
  só alcança as caixas dela e as compartilhadas com ela) e prefira o certificado ao segredo do
  cliente (a chave privada fica só no servidor do CLEAN); no SharePoint, prefira
  `Sites.Selected` com os sites liberados um a um (`Files.Read.All` dá acesso a todos os arquivos do
  locatário); só conceda permissões de escrita e marque *Permitir exclusão* onde a exclusão for
  realmente usada;
- com a exclusão permitida, qualquer pessoa com acesso à interface pode excluir itens pelo
  relatório: defina `AUTH_USER`/`AUTH_PASSWORD` (o registro mostra essa conta, que é única, e o
  endereço de acesso de quem excluiu);
- a interface tem proteção contra CSRF e política de segurança de conteúdo. Sem a exclusão
  (*Somente analisar* ou *Permitir exclusão* desmarcado), o CLEAN apenas lê os arquivos e as
  mensagens, sem alterá-los.

## Formatos suportados e limitações

| Tipo | Extensões | Observação |
|---|---|---|
| Word | .docx, .docm, .dotx, .doc | Inclui cabeçalhos, rodapés, notas e comentários. |
| Excel | .xlsx, .xlsm, .xltx, .xls | Informa a planilha e a linha; números (CPF gravado como número) também são lidos. |
| PowerPoint | .pptx, .pptm, .ppsx, .ppt | Informa o slide (.pptx); inclui anotações. |
| PDF | .pdf | Informa a página. |
| OpenDocument | .odt, .ods, .odp | |
| E-mails e páginas | .msg, .eml, .mht/.mhtml | Assunto, remetente, destinatários, corpo, nomes e conteúdo dos anexos (inclusive mensagens anexadas). |
| Outros | .rtf, .htm/.html, .txt, .csv, .log, .xml, .json e demais textos | Codificação detectada automaticamente. |
| Compactados | .zip | Apenas os nomes dos arquivos internos (inclusive os criados pelo Explorer do Windows). |

Limitações conhecidas:

- imagens e PDFs digitalizados não são lidos (não há OCR);
- o conteúdo de arquivos dentro de .zip, .7z, .rar e de caixas de correio .pst/.ost não é analisado;
- mensagens criptografadas (S/MIME, PGP) ou protegidas pelo Microsoft Purview (IRM) têm apenas o
  assunto e os remetentes verificados; anexos `winmail.dat` (TNEF) aparecem só pelo nome;
- arquivos protegidos por senha têm apenas o nome verificado (aparecem como "Protegido por senha");
- arquivos danificados ou que demoram demais para ser lidos aparecem na aba **Erros** e a análise
  continua; expressões regulares muito lentas são interrompidas por arquivo;
- Word 6/95 e Excel 5/95 têm extração aproximada;
- a auditoria só cobre o período mantido no log de Segurança e precisa estar habilitada antes;
- datas e números exportados usam o fuso horário e o formato do servidor.

## Desenvolvimento

```bash
npm install
npm test          # testes automatizados (node:test)
npm run dev       # servidor com recarga automática
```

Estrutura:

```
src/
  server.js, app.js, config.js, store.js   servidor, rotas e persistência (JSON/NDJSON)
  secrets.js                               cifragem das credenciais (caixas de e-mail, OneDrive e SharePoint)
  routes/                                  API REST (/api/repositories, /api/lists, /api/mail-sources, /api/scans, /api/schedules)
  schedule/
    recurrence.js                          regras de recorrência: validação, próximas execuções e descrição
    scheduler.js                           agendador: horários, sobreposição, horários perdidos, período e retenção
  cloud/
    microsoft-auth.js                      OAuth 2.0 da Microsoft: segredo, certificado (asserção PS256), conta conectada (código de dispositivo e renovação)
    certificate.js                         certificado autoassinado (X.509) e importação em PEM
    pending-credentials.js                 entradas de contas e certificados ainda não salvos (memória)
    graph-client.js                        cliente do Microsoft Graph (token, novas tentativas, paginação)
    drives.js                              OneDrive e SharePoint: contas, sites, bibliotecas, arquivos, exclusão
  scan/
    scanner.js, worker.js, manager.js      análise em worker thread, fila e cancelamento
    delete.js                              exclusão dos arquivos encontrados
    walker.js                              percurso das pastas e exclusões
    matcher.js, presets.js                 busca de termos (Aho-Corasick), regex e validadores
    owner.js, audit.js, powershell.js      proprietário NTFS e log de auditoria via PowerShell
    extractors/                            leitura de cada formato de arquivo (e das mensagens MIME e anexos)
  mail/
    scanner.js                             análise das caixas, o raio-x e as listagens (contas e mensagens), na mesma worker thread
    graph.js, gmail.js, imap.js            conectores Microsoft 365, Google Workspace e IMAP (varredura, catálogo de contas e exclusão)
    http.js, common.js                     requisições com novas tentativas, pastas e caixas ignoradas
  types/
    catalog.js                             busca por tipo: categorias, extensões e validação
    signature.js                           tipo real pelos primeiros bytes (arquivos renomeados)
  retention/                               políticas de retenção: critérios de data e limites
  report/                                  filtros, resumo e exportações (xlsx, csv, html)
public/                                    interface web (HTML, CSS e JavaScript, sem build)
test/                                      testes e arquivos de exemplo (fixtures)
scripts/criar-dados-demo.js                dados de demonstração
```

Os testes dos scripts PowerShell usam versões simuladas de `Get-Acl` e `Get-WinEvent` e rodam no
Windows ou onde houver PowerShell 7 (`pwsh`); sem PowerShell, são ignorados. Os testes de e-mail usam
servidores simulados do Microsoft Entra ID (segredo, certificado, código de dispositivo e renovação
dos tokens), do Microsoft Graph (inclusive OneDrive e SharePoint e o acesso delegado), das APIs do
Google e de IMAP (com XOAUTH2) (`test/helpers`), sem acesso à internet.
