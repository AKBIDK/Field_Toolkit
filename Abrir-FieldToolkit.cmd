@echo off
setlocal DisableDelayedExpansion
set "FT_SCRIPT=%~dp0FieldToolkit.ps1"
set "FT_MODE=normal"
if /I "%~1"=="-Demo" set "FT_MODE=demo"
if /I "%~1"=="-TesteRelatorio" set "FT_MODE=test"
if not exist "%FT_SCRIPT%" (
    echo FieldToolkit.ps1 nao encontrado. Extraia todos os arquivos do ZIP na mesma pasta.
    pause
    exit /b 1
)
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -Command "& { $ErrorActionPreference = 'Stop'; try { $scriptFile = $env:FT_SCRIPT; $policy = Get-ExecutionPolicy; if ($policy -in 'Restricted','AllSigned') { throw ('A politica efetiva do Windows e ' + $policy + '. Consulte o administrador do computador.') }; if (Get-Item -LiteralPath $scriptFile -Stream Zone.Identifier -ErrorAction SilentlyContinue) { $answer = Read-Host 'O Windows marcou FieldToolkit.ps1 como baixado. Se voce confia neste arquivo, digite S para desbloquear somente este script'; if ($answer -ne 'S') { return }; Unblock-File -LiteralPath $scriptFile }; switch ($env:FT_MODE) { 'demo' { & $scriptFile -Demo }; 'test' { & $scriptFile -TesteRelatorio }; default { & $scriptFile } } } catch { Write-Host $_.Exception.Message -ForegroundColor Red; exit 1 } }"
if errorlevel 1 (
    pause
    exit /b 1
)
if /I "%FT_MODE%"=="test" pause
