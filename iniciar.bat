@echo off
rem Inicia o CLEAN (servidor web). Mantenha esta janela aberta enquanto usar a aplicacao.
setlocal
cd /d "%~dp0"

rem Se este servidor acessa a internet por um proxy (analise de e-mail do Microsoft 365 ou do
rem Google Workspace), retire o "rem" das duas linhas abaixo e informe o endereco do proxy:
rem set NODE_USE_ENV_PROXY=1
rem set HTTPS_PROXY=http://proxy.empresa.local:3128

where node >NUL 2>NUL
if errorlevel 1 (
  echo Node.js nao encontrado. Instale a versao LTS em https://nodejs.org e execute novamente.
  pause
  exit /b 1
)

rem Instala as dependencias na primeira execucao e depois de atualizacoes que incluam novas.
node -e "const fs=require('fs');const d=Object.keys(require('./package.json').dependencies||{});process.exit(d.every((x)=>fs.existsSync('node_modules/'+x))?0:1)"
if errorlevel 1 (
  echo Instalando dependencias...
  call npm install --omit=dev
  if errorlevel 1 (
    echo Falha ao instalar as dependencias.
    pause
    exit /b 1
  )
)

node src\server.js
pause
