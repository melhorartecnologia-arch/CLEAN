@echo off
rem Inicia o CLEAN (servidor web). Mantenha esta janela aberta enquanto usar a aplicacao.
setlocal
cd /d "%~dp0"

where node >NUL 2>NUL
if errorlevel 1 (
  echo Node.js nao encontrado. Instale a versao LTS em https://nodejs.org e execute novamente.
  pause
  exit /b 1
)

if not exist node_modules (
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
