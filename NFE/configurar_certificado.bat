@echo off
title Configuracao de Certificado Digital - ACBrLibMDFe
echo ===========================================
echo     CONFIGURACAO DE CERTIFICADO DIGITAL
echo           Sistema MDFe - ACBrLibMDFe
echo ===========================================

echo.
echo Este script ira ajudar voce a configurar o certificado digital
echo no arquivo ACBrLibMDFe.ini do seu sistema.
echo.

REM Verificar se o arquivo INI existe
if not exist "ACBrLibMDFe.ini" (
    echo ERRO: Arquivo ACBrLibMDFe.ini nao encontrado!
    echo Execute primeiro o script instalar_acbr_producao.bat
    pause
    exit /b 1
)

echo Configuracoes atuais:
echo =====================

REM Mostrar configuracoes atuais
findstr /C:"Arquivo=" ACBrLibMDFe.ini
findstr /C:"Senha=" ACBrLibMDFe.ini  
findstr /C:"Ambiente=" ACBrLibMDFe.ini

echo.
echo Opcoes de configuracao:
echo.
echo 1. Configurar certificado A1 (arquivo .pfx/.p12)
echo 2. Configurar certificado A3 (numero de serie)
echo 3. Alterar ambiente (Producao/Homologacao)
echo 4. Testar configuracao atual
echo 5. Mostrar configuracao completa
echo 6. Sair
echo.

set /p opcao="Escolha uma opcao (1-6): "

if "%opcao%"=="1" goto ConfigA1
if "%opcao%"=="2" goto ConfigA3
if "%opcao%"=="3" goto ConfigAmbiente
if "%opcao%"=="4" goto TestarConfig
if "%opcao%"=="5" goto MostrarConfig
if "%opcao%"=="6" goto Sair
goto Inicio

:ConfigA1
echo.
echo ========================================
echo  CONFIGURACAO CERTIFICADO A1 (ARQUIVO)
echo ========================================
echo.

REM Solicitar caminho do certificado
echo Digite o caminho completo do arquivo do certificado (.pfx ou .p12):
echo Exemplo: C:\Certificados\MeuCertificado.pfx
echo.
set /p caminho_cert="Caminho do certificado: "

REM Verificar se arquivo existe
if not exist "%caminho_cert%" (
    echo.
    echo ERRO: Arquivo nao encontrado: %caminho_cert%
    echo Verifique o caminho e tente novamente.
    pause
    goto ConfigA1
)

REM Solicitar senha
echo.
set /p senha_cert="Digite a senha do certificado: "

REM Fazer backup do INI atual
copy "ACBrLibMDFe.ini" "ACBrLibMDFe.ini.bak" >nul

REM Criar script temporario para substituir as configuracoes
echo @echo off > temp_config.bat
echo powershell -Command "(Get-Content 'ACBrLibMDFe.ini') -replace '^Arquivo=.*', 'Arquivo=%caminho_cert%' | Set-Content 'ACBrLibMDFe.ini'" >> temp_config.bat
echo powershell -Command "(Get-Content 'ACBrLibMDFe.ini') -replace '^Senha=.*', 'Senha=%senha_cert%' | Set-Content 'ACBrLibMDFe.ini'" >> temp_config.bat
echo powershell -Command "(Get-Content 'ACBrLibMDFe.ini') -replace '^NumeroSerie=.*', 'NumeroSerie=' | Set-Content 'ACBrLibMDFe.ini'" >> temp_config.bat

REM Executar configuracao
call temp_config.bat
del temp_config.bat

echo.
echo ✓ Certificado A1 configurado com sucesso!
echo   Arquivo: %caminho_cert%
echo   Backup salvo em: ACBrLibMDFe.ini.bak
echo.
pause
goto Inicio

:ConfigA3
echo.
echo ========================================
echo  CONFIGURACAO CERTIFICADO A3 (SERIE)
echo ========================================
echo.
echo Para usar certificado A3, voce precisa do numero de serie.
echo.
echo Para descobrir o numero de serie:
echo 1. Abra o Internet Explorer
echo 2. Va em Ferramentas > Opcoes da Internet > Conteudo > Certificados
echo 3. Encontre seu certificado e veja o campo "Numero de Serie"
echo.
set /p num_serie="Digite o numero de serie (sem espacos): "

if "%num_serie%"=="" (
    echo ERRO: Numero de serie nao pode estar vazio!
    pause
    goto ConfigA3
)

REM Fazer backup do INI atual
copy "ACBrLibMDFe.ini" "ACBrLibMDFe.ini.bak" >nul

REM Criar script temporario para substituir as configuracoes
echo @echo off > temp_config.bat
echo powershell -Command "(Get-Content 'ACBrLibMDFe.ini') -replace '^Arquivo=.*', 'Arquivo=' | Set-Content 'ACBrLibMDFe.ini'" >> temp_config.bat
echo powershell -Command "(Get-Content 'ACBrLibMDFe.ini') -replace '^Senha=.*', 'Senha=' | Set-Content 'ACBrLibMDFe.ini'" >> temp_config.bat
echo powershell -Command "(Get-Content 'ACBrLibMDFe.ini') -replace '^NumeroSerie=.*', 'NumeroSerie=%num_serie%' | Set-Content 'ACBrLibMDFe.ini'" >> temp_config.bat

REM Executar configuracao
call temp_config.bat
del temp_config.bat

echo.
echo ✓ Certificado A3 configurado com sucesso!
echo   Numero de Serie: %num_serie%
echo   Backup salvo em: ACBrLibMDFe.ini.bak
echo.
pause
goto Inicio

:ConfigAmbiente
echo.
echo ========================================
echo      CONFIGURACAO DE AMBIENTE
echo ========================================
echo.
echo Ambiente atual:
findstr /C:"Ambiente=" ACBrLibMDFe.ini | head -2

echo.
echo Opcoes:
echo 1. Producao (ambiente real)
echo 2. Homologacao (ambiente de testes)
echo.
set /p amb_opcao="Escolha o ambiente (1 ou 2): "

if "%amb_opcao%"=="1" (
    set novo_ambiente=1
    set nome_ambiente=PRODUCAO
) else if "%amb_opcao%"=="2" (
    set novo_ambiente=2  
    set nome_ambiente=HOMOLOGACAO
) else (
    echo Opcao invalida!
    pause
    goto ConfigAmbiente
)

REM Fazer backup
copy "ACBrLibMDFe.ini" "ACBrLibMDFe.ini.bak" >nul

REM Alterar ambiente nas duas secoes
powershell -Command "(Get-Content 'ACBrLibMDFe.ini') -replace '^Ambiente=.*', 'Ambiente=%novo_ambiente%' | Set-Content 'ACBrLibMDFe.ini'"

echo.
echo ✓ Ambiente alterado para: %nome_ambiente%
echo   Backup salvo em: ACBrLibMDFe.ini.bak
echo.
pause
goto Inicio

:TestarConfig
echo.
echo ========================================
echo        TESTANDO CONFIGURACAO
echo ========================================
echo.

REM Verificar se DLL existe
if exist "ACBrMDFe32.dll" (
    echo ✓ DLL ACBrMDFe32.dll encontrada
) else if exist "ACBrMDFe64.dll" (
    echo ✓ DLL ACBrMDFe64.dll encontrada  
) else (
    echo ✗ DLL ACBrMDFe nao encontrada!
    echo   Execute o script de instalacao primeiro.
    pause
    goto Inicio
)

REM Mostrar configuracoes atuais de certificado
echo.
echo Configuracoes do certificado:
findstr /C:"Arquivo=" ACBrLibMDFe.ini
findstr /C:"NumeroSerie=" ACBrLibMDFe.ini
findstr /C:"Ambiente=" ACBrLibMDFe.ini

REM Verificar se certificado foi configurado
findstr /C:"Arquivo=" ACBrLibMDFe.ini | findstr /V /C:"Arquivo=$" >nul
if %errorlevel% equ 0 (
    echo ✓ Certificado A1 configurado
) else (
    findstr /C:"NumeroSerie=" ACBrLibMDFe.ini | findstr /V /C:"NumeroSerie=$" >nul
    if %errorlevel% equ 0 (
        echo ✓ Certificado A3 configurado
    ) else (
        echo ⚠ ATENCAO: Nenhum certificado configurado!
        echo   Configure um certificado antes de usar o sistema.
    )
)

echo.
echo Para teste completo, execute TesteACBrProducao.vbs
echo ou use o modulo TesteACBrProducao.bas no VB6.
echo.
pause
goto Inicio

:MostrarConfig
echo.
echo ========================================
echo     CONFIGURACAO COMPLETA ATUAL
echo ========================================
echo.
type "ACBrLibMDFe.ini"
echo.
pause
goto Inicio

:Sair
echo.
echo Configuracao finalizada.
echo.
echo IMPORTANTE:
echo - Teste sempre em homologacao antes de usar em producao
echo - Mantenha backup dos arquivos de configuracao
echo - Verifique se o certificado esta dentro do prazo de validade
echo.
pause
exit

:Inicio
cls
echo ===========================================
echo     CONFIGURACAO DE CERTIFICADO DIGITAL
echo           Sistema MDFe - ACBrLibMDFe  
echo ===========================================
goto Menu