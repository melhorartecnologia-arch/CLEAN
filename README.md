# CLEAN — Análise de repositórios de arquivos

Aplicação web em Node.js que percorre repositórios de arquivos do Windows (pastas locais e
compartilhamentos de rede), procura os termos de uma **lista de referência** no **nome** e no
**conteúdo** dos arquivos e gera relatórios com **a informação encontrada** e **o último usuário que
interagiu com cada arquivo**.

- Listas de referência com termos de texto (sem diferenciar maiúsculas e acentos: "salario" encontra
  "SALÁRIO") ou expressões regulares, com modelos prontos e validados: CPF, CNPJ (inclusive o novo
  CNPJ alfanumérico), PIS/PASEP, e-mail, telefone, cartão de crédito e "senha em texto".
- Lê o conteúdo de Word, Excel e PowerPoint (novos e 97-2003), PDF, OpenDocument (LibreOffice),
  RTF, e-mails do Outlook (.msg), HTML, textos e CSV (UTF-8, UTF-16 e Windows-1252) e os nomes dos
  arquivos dentro de .zip. Arquivos protegidos por senha são identificados.
- Último usuário a partir de três fontes, da mais precisa para a menos precisa: **log de auditoria
  do Windows** (quem acessou ou alterou por último), **metadados do documento** ("salvo por último
  por") e **proprietário do arquivo (NTFS)**.
- Relatório na tela com filtros, gráficos e trechos em que cada termo aparece (com página, planilha,
  slide ou linha), e exportação para **Excel**, **CSV**, **HTML** (para imprimir) e **JSON**.
- Sem banco de dados e sem etapa de compilação: basta instalar o Node.js e executar.

## Sumário

1. [Requisitos](#requisitos)
2. [Instalação e primeiro uso](#instalação-e-primeiro-uso)
3. [Como usar](#como-usar)
4. [Como o último usuário é identificado](#como-o-último-usuário-é-identificado)
5. [Habilitando a auditoria do Windows (opcional)](#habilitando-a-auditoria-do-windows-opcional)
6. [Conta de serviço e permissões](#conta-de-serviço-e-permissões)
7. [Executando como serviço](#executando-como-serviço)
8. [Configuração](#configuração)
9. [Segurança](#segurança)
10. [Formatos suportados e limitações](#formatos-suportados-e-limitações)
11. [Desenvolvimento](#desenvolvimento)

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
2. **Listas de referência** — crie uma lista e adicione termos:
   - **Texto**: ignora maiúsculas, acentos e espaços repetidos. Marque *Palavra inteira* para que
     "ana" não encontre "banana".
   - **Expressão regular** (sintaxe JavaScript), opcionalmente com validação (CPF, CNPJ, PIS,
     cartão) para descartar números inválidos, e um nome para o relatório.
   - *Adicionar vários*, *Importar .txt/.csv* (um termo por linha ou a primeira coluna do CSV) e
     *Adicionar modelo pronto*.
   - *Testar os termos* mostra o que seria encontrado em um texto de exemplo antes de salvar.
3. **Nova análise** — escolha repositórios e listas e o que verificar:
   - nome do arquivo (ou o caminho completo, incluindo os nomes das pastas) e/ou conteúdo;
   - somente arquivos modificados a partir de uma data;
   - tamanho máximo para ler o conteúdo (arquivos maiores têm apenas o nome verificado; textos
     longos, só o início);
   - quantidade de arquivos processados em paralelo.
4. **Relatório** — acompanha o progresso em tempo real. Mostra os números da análise, os termos
   mais encontrados, os últimos usuários (clique numa barra para filtrar) e a lista de arquivos com
   ocorrências. Cada arquivo abre os detalhes: caminho, datas, todas as fontes de "último usuário" e
   os trechos com o termo destacado. As exportações respeitam os filtros aplicados:
   - **Excel**: abas *Resumo*, *Arquivos* (um arquivo por linha), *Ocorrências* (um termo por
     linha, com valores e trechos) e *Erros*;
   - **CSV**: uma linha por arquivo e termo, no padrão do Excel em português (`;` e UTF-8);
   - **HTML**: relatório para leitura ou impressão; **JSON**: para integração com outros sistemas.

As análises rodam em segundo plano (é possível fechar o navegador) e podem ser canceladas a qualquer
momento; o que já foi encontrado é mantido.

## Como o último usuário é identificado

O NTFS não registra "quem alterou por último". O CLEAN combina as fontes disponíveis e informa no
relatório qual foi usada:

| Prioridade | Fonte | O que informa | Observações |
|---|---|---|---|
| 1 | **Log de auditoria** (eventos 4663 e 5145) | Quem acessou por último (e a ação: leitura, gravação, exclusão, permissões) e quem alterou por último | É a fonte mais fiel, mas exige a auditoria habilitada **antes** (ver abaixo) e vale apenas para o período mantido no log. |
| 2 | **Metadados do documento** | "Salvo por último por" do Word, Excel, PowerPoint, LibreOffice, RTF e .msg | É o nome configurado no Office de quem salvou (normalmente o nome completo). Não muda quando o arquivo é alterado por outros programas. |
| 3 | **Proprietário (NTFS)** | Dono do arquivo | Em geral quem **criou** o arquivo (às vezes o grupo Administradores). Usado quando não há as fontes acima. |

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

- permissão de **leitura** (NTFS e compartilhamento) nas pastas analisadas;
- participação no grupo **Leitores de Log de Eventos** (*Event Log Readers*) de cada servidor de
  arquivos, se usar a auditoria; para consultar outro computador, a regra de firewall
  *Gerenciamento Remoto do Log de Eventos* deve estar habilitada nele.

Pastas que a conta não consegue abrir aparecem na aba **Erros** do relatório.

## Executando como serviço

Para que o CLEAN inicie com o Windows, sem sessão aberta, use o Agendador de Tarefas (nativo). Em
um PowerShell como administrador:

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
| `AUTH_USER` / `AUTH_PASSWORD` | — | Ativa usuário e senha (autenticação HTTP básica) para a interface e a API. |
| `DATA_DIR` | `data` | Pasta com a configuração (`db.json`) e os resultados de cada análise (`scans\<id>`). |
| `MAX_CONCURRENT_SCANS` | `1` | Análises simultâneas; as demais aguardam na fila. |
| `POWERSHELL_PATH` | `powershell.exe` | PowerShell usado para o proprietário e o log de auditoria. |

## Segurança

Os relatórios mostram nomes de arquivos, usuários e **trechos do conteúdo** — incluindo, conforme a
lista, dados pessoais e senhas encontradas. Por isso:

- por padrão o servidor só aceita conexões do próprio computador (`HOST=127.0.0.1`);
- ao liberar o acesso pela rede, defina `AUTH_USER` e `AUTH_PASSWORD` e, de preferência, publique o
  CLEAN atrás de um proxy HTTPS (IIS com URL Rewrite/ARR, por exemplo);
- proteja a pasta `data` com permissões NTFS restritas (ela guarda os resultados das análises);
- a interface tem proteção contra CSRF e política de segurança de conteúdo; o CLEAN nunca altera
  os arquivos analisados, apenas os lê.

## Formatos suportados e limitações

| Tipo | Extensões | Observação |
|---|---|---|
| Word | .docx, .docm, .dotx, .doc | Inclui cabeçalhos, rodapés, notas e comentários. |
| Excel | .xlsx, .xlsm, .xltx, .xls | Informa a planilha e a linha; números (CPF gravado como número) também são lidos. |
| PowerPoint | .pptx, .pptm, .ppsx, .ppt | Informa o slide (.pptx); inclui anotações. |
| PDF | .pdf | Informa a página. |
| OpenDocument | .odt, .ods, .odp | |
| Outros | .rtf, .msg, .htm/.html, .txt, .csv, .log, .xml, .json e demais textos | Codificação detectada automaticamente. |
| Compactados | .zip | Apenas os nomes dos arquivos internos. |

Limitações conhecidas:

- imagens e PDFs digitalizados não são lidos (não há OCR);
- o conteúdo de arquivos dentro de .zip, .7z, .rar e de caixas de correio .pst/.ost não é analisado;
- arquivos protegidos por senha têm apenas o nome verificado (aparecem como "Protegido por senha");
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
  routes/                                  API REST (/api/repositories, /api/lists, /api/scans)
  scan/
    scanner.js, worker.js, manager.js      análise em worker thread, fila e cancelamento
    walker.js                              percurso das pastas e exclusões
    matcher.js, presets.js                 busca de termos (Aho-Corasick), regex e validadores
    owner.js, audit.js, powershell.js      proprietário NTFS e log de auditoria via PowerShell
    extractors/                            leitura de cada formato de arquivo
  report/                                  filtros, resumo e exportações (xlsx, csv, html)
public/                                    interface web (HTML, CSS e JavaScript, sem build)
test/                                      testes e arquivos de exemplo (fixtures)
scripts/criar-dados-demo.js                dados de demonstração
```

Os testes dos scripts PowerShell usam versões simuladas de `Get-Acl` e `Get-WinEvent` e rodam no
Windows ou onde houver PowerShell 7 (`pwsh`); sem PowerShell, são ignorados.
