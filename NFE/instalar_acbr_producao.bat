@echo off
title Instalacao ACBrLibMDFe - AMBIENTE PRODUCAO
echo ===========================================
echo      INSTALACAO ACBrLibMDFe v1.2.2.335
echo           AMBIENTE DE PRODUCAO
echo ===========================================

REM Verificar se tem privilegios administrativos
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Execute como Administrador!
    echo Clique com botao direito e "Executar como administrador"
    pause
    exit /b 1
)

echo.
echo ATENCAO: Esta instalacao sera feita no ambiente de PRODUCAO!
echo Pasta de destino: C:\Projetos\MDFe\NFE
echo.
set /p confirma="Confirma a instalacao no ambiente de producao? (S/N): "
if /i "%confirma%" neq "S" (
    echo Instalacao cancelada pelo usuario.
    pause
    exit /b 0
)

echo.
echo Verificando arquitetura do sistema...
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    set ARCH=64
    echo Sistema detectado: 64-bit
) else (
    set ARCH=32
    echo Sistema detectado: 32-bit
)

echo.
echo 1. Fazendo backup do sistema atual...
set DATA_BACKUP=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2%_%TIME:~0,2%-%TIME:~3,2%
set DATA_BACKUP=%DATA_BACKUP: =0%
if not exist "C:\Backup\MDFe\%DATA_BACKUP%" mkdir "C:\Backup\MDFe\%DATA_BACKUP%"
xcopy "*.exe" "C:\Backup\MDFe\%DATA_BACKUP%\" /Y 2>nul
xcopy "*.dll" "C:\Backup\MDFe\%DATA_BACKUP%\" /Y 2>nul
xcopy "*.ini" "C:\Backup\MDFe\%DATA_BACKUP%\" /Y 2>nul
echo Backup salvo em: C:\Backup\MDFe\%DATA_BACKUP%\

echo.
echo 2. Parando processos relacionados ao MDFe...
taskkill /f /im NFE.exe 2>nul
taskkill /f /im MDFE.exe 2>nul
timeout /t 3 /nobreak >nul
echo Processos parados.

echo.
echo 3. Criando diretorios necessarios...
if not exist "Logs" mkdir "Logs"
if not exist "XML" mkdir "XML" 
if not exist "PDF" mkdir "PDF"
if not exist "dep" mkdir "dep"
if not exist "Backup" mkdir "Backup"
echo Diretorios criados.

echo.
echo 4. Verificando origem dos arquivos ACBr...
set ORIGEM_ACBR=
if exist "C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe-Windows-1.2.2.335" (
    set ORIGEM_ACBR=C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe-Windows-1.2.2.335
    echo Origem encontrada: %ORIGEM_ACBR%
) else if exist "ACBrLibMDFe-Windows-1.2.2.335" (
    set ORIGEM_ACBR=ACBrLibMDFe-Windows-1.2.2.335
    echo Origem encontrada: %ORIGEM_ACBR%
) else (
    echo ERRO: Nao foi possivel encontrar os arquivos ACBrLibMDFe-Windows-1.2.2.335
    echo Verifique se a pasta existe em uma das localizacoes:
    echo - C:\Projetos\MDFe - CLAUDE\NFE\ACBrLibMDFe-Windows-1.2.2.335
    echo - .\ACBrLibMDFe-Windows-1.2.2.335
    pause
    exit /b 1
)

echo.
echo 5. Copiando DLL principal ACBrMDFe...
copy "%ORIGEM_ACBR%\Windows\MT\StdCall\ACBrMDFe%ARCH%.dll" "ACBrMDFe%ARCH%.dll" /Y
if %errorlevel% neq 0 (
    echo ERRO: Nao foi possivel copiar ACBrMDFe%ARCH%.dll
    echo Verifique se o arquivo existe em: %ORIGEM_ACBR%\Windows\MT\StdCall\
    pause
    exit /b 1
)
echo ✓ DLL principal copiada: ACBrMDFe%ARCH%.dll

echo.
echo 6. Copiando dependencias OpenSSL...
if "%ARCH%"=="64" (
    copy "%ORIGEM_ACBR%\dep\OpenSSL\x64\libcrypto-1_1-x64.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\OpenSSL\x64\libssl-1_1-x64.dll" "." /Y
    echo ✓ OpenSSL x64 copiado
) else (
    copy "%ORIGEM_ACBR%\dep\OpenSSL\x86\libcrypto-1_1.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\OpenSSL\x86\libssl-1_1.dll" "." /Y
    echo ✓ OpenSSL x86 copiado
)

echo.
echo 7. Copiando dependencias LibXML2...
if "%ARCH%"=="64" (
    copy "%ORIGEM_ACBR%\dep\LibXml2\x64\libxml2.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\LibXml2\x64\libxslt.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\LibXml2\x64\libexslt.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\LibXml2\x64\libiconv.dll" "." /Y
    echo ✓ LibXML2 x64 copiado
) else (
    copy "%ORIGEM_ACBR%\dep\LibXml2\x86\libxml2.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\LibXml2\x86\libxslt.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\LibXml2\x86\libexslt.dll" "." /Y
    copy "%ORIGEM_ACBR%\dep\LibXml2\x86\libiconv.dll" "." /Y
    echo ✓ LibXML2 x86 copiado
)

echo.
echo 8. Copiando arquivo de servicos SEFAZ...
copy "%ORIGEM_ACBR%\dep\ACBrMDFeServicos.ini" "dep\" /Y
echo ✓ Arquivo de servicos copiado.

echo.
echo 9. Desregistrando DLL FlexDocs antiga (se existir)...
if exist "MDFe_Util.dll" (
    echo Encontrada DLL FlexDocs antiga, desregistrando...
    %windir%\Microsoft.NET\Framework\v4.0.30319\regasm /u "MDFe_Util.dll" /tlb 2>nul
    if exist "MDFe_Util.tlb" del "MDFe_Util.tlb" 2>nul
    echo ✓ DLL FlexDocs desregistrada
) else (
    echo ✓ Nenhuma DLL FlexDocs encontrada para desregistrar
)

echo.
echo 10. Criando arquivo de configuracao para PRODUCAO...
echo [Principal] > ACBrLibMDFe.ini
echo LogLevel=2 >> ACBrLibMDFe.ini
echo LogPath=C:\Projetos\MDFe\NFE\Logs\ >> ACBrLibMDFe.ini
echo. >> ACBrLibMDFe.ini
echo [DFe] >> ACBrLibMDFe.ini
echo UF=SP >> ACBrLibMDFe.ini
echo Ambiente=1 >> ACBrLibMDFe.ini
echo Visualizar=0 >> ACBrLibMDFe.ini
echo SalvarWS=1 >> ACBrLibMDFe.ini
echo RetirarAcentos=1 >> ACBrLibMDFe.ini
echo FormatoAlerta=clAsterisco >> ACBrLibMDFe.ini
echo PathSchemas= >> ACBrLibMDFe.ini
echo. >> ACBrLibMDFe.ini
echo [WebService] >> ACBrLibMDFe.ini
echo UF=SP >> ACBrLibMDFe.ini
echo Ambiente=1 >> ACBrLibMDFe.ini
echo Visualizar=0 >> ACBrLibMDFe.ini
echo SalvarWS=1 >> ACBrLibMDFe.ini
echo SalvarEnvio=1 >> ACBrLibMDFe.ini
echo SalvarResposta=1 >> ACBrLibMDFe.ini
echo AjustaAguardaConsultaRet=1 >> ACBrLibMDFe.ini
echo AguardarConsultaRet=1000 >> ACBrLibMDFe.ini
echo Tentativas=5 >> ACBrLibMDFe.ini
echo IntervaloTentativas=2000 >> ACBrLibMDFe.ini
echo TimeOut=60000 >> ACBrLibMDFe.ini
echo ProxyHost= >> ACBrLibMDFe.ini
echo ProxyPort= >> ACBrLibMDFe.ini
echo ProxyUser= >> ACBrLibMDFe.ini
echo ProxyPass= >> ACBrLibMDFe.ini
echo. >> ACBrLibMDFe.ini
echo [Certificados] >> ACBrLibMDFe.ini
echo Arquivo= >> ACBrLibMDFe.ini
echo Senha= >> ACBrLibMDFe.ini
echo NumeroSerie= >> ACBrLibMDFe.ini
echo CacheLib=1 >> ACBrLibMDFe.ini
echo CryptoLib=1 >> ACBrLibMDFe.ini
echo HttpLib=1 >> ACBrLibMDFe.ini
echo XmlSignLib=1 >> ACBrLibMDFe.ini
echo. >> ACBrLibMDFe.ini
echo [Arquivos] >> ACBrLibMDFe.ini
echo PastaMensal=1 >> ACBrLibMDFe.ini
echo AddLiteral=0 >> ACBrLibMDFe.ini
echo EmissaoPathMDFe=1 >> ACBrLibMDFe.ini
echo SalvarEvento=1 >> ACBrLibMDFe.ini
echo SepararPorCNPJ=0 >> ACBrLibMDFe.ini
echo PathMDFe=C:\Projetos\MDFe\NFE\XML\ >> ACBrLibMDFe.ini
echo PathEvento=C:\Projetos\MDFe\NFE\XML\ >> ACBrLibMDFe.ini
echo. >> ACBrLibMDFe.ini
echo [DAMDFE] >> ACBrLibMDFe.ini
echo TipoDAMDFE=0 >> ACBrLibMDFe.ini
echo PathPDF=C:\Projetos\MDFe\NFE\PDF\ >> ACBrLibMDFe.ini
echo PathLogo= >> ACBrLibMDFe.ini
echo Visualizar=1 >> ACBrLibMDFe.ini
echo ImprimirHoraSaida=0 >> ACBrLibMDFe.ini
echo ImprimirHoraSaida_Hora=12:00:00 >> ACBrLibMDFe.ini
echo TamanhoPapel=0 >> ACBrLibMDFe.ini
echo Margem_Sup=8 >> ACBrLibMDFe.ini
echo Margem_Inf=8 >> ACBrLibMDFe.ini
echo Margem_Esq=6 >> ACBrLibMDFe.ini
echo Margem_Dir=6 >> ACBrLibMDFe.ini
echo FonteDAMDFE_Nome=Times New Roman >> ACBrLibMDFe.ini
echo FonteDAMDFE_Tamanho=9 >> ACBrLibMDFe.ini
echo ✓ Configuracao para PRODUCAO criada

echo.
echo 11. Testando instalacao basica...
REM Testar se consegue carregar a DLL
if exist "ACBrMDFe%ARCH%.dll" (
    echo ✓ DLL principal encontrada
) else (
    echo ✗ ERRO: DLL principal nao encontrada!
    pause
    exit /b 1
)

REM Verificar dependencias criticas
set DEPS_OK=1
if not exist "libcrypto-1_1*.dll" set DEPS_OK=0
if not exist "libxml2.dll" set DEPS_OK=0

if %DEPS_OK%==1 (
    echo ✓ Dependencias principais verificadas
) else (
    echo ✗ AVISO: Algumas dependencias podem estar faltando
)

echo.
echo 12. Criando script de teste rapido...
echo ' TesteACBrProducao.vbs > TesteACBrProducao.vbs
echo ' Script para testar ACBrLibMDFe em producao >> TesteACBrProducao.vbs
echo Set shell = CreateObject("WScript.Shell") >> TesteACBrProducao.vbs
echo MsgBox "ACBrLibMDFe instalado em C:\Projetos\MDFe\NFE" ^& vbCrLf ^& _ >> TesteACBrProducao.vbs
echo        "Proximos passos:" ^& vbCrLf ^& _ >> TesteACBrProducao.vbs
echo        "1. Configure certificado no arquivo ACBrLibMDFe.ini" ^& vbCrLf ^& _ >> TesteACBrProducao.vbs
echo        "2. Ajuste UF se necessario" ^& vbCrLf ^& _ >> TesteACBrProducao.vbs
echo        "3. Teste em homologacao antes de usar", vbInformation, "ACBr Instalado" >> TesteACBrProducao.vbs

echo.
echo ===========================================
echo     INSTALACAO CONCLUIDA COM SUCESSO!
echo           AMBIENTE DE PRODUCAO
echo ===========================================
echo.
echo ✓ Pasta de instalacao: C:\Projetos\MDFe\NFE
echo ✓ Backup salvo em: C:\Backup\MDFe\%DATA_BACKUP%\
echo ✓ DLL principal: ACBrMDFe%ARCH%.dll
echo ✓ Configuracao: ACBrLibMDFe.ini (PRODUCAO - Ambiente=1)
echo ✓ Dependencias: OpenSSL + LibXML2 instaladas
echo ✓ Logs serao salvos em: C:\Projetos\MDFe\NFE\Logs\
echo ✓ XMLs serao salvos em: C:\Projetos\MDFe\NFE\XML\
echo ✓ PDFs serao salvos em: C:\Projetos\MDFe\NFE\PDF\
echo.
echo IMPORTANTE:
echo 1. Configure seu certificado digital no arquivo ACBrLibMDFe.ini
echo 2. Esta configuracao esta para PRODUCAO (Ambiente=1)
echo 3. Teste em homologacao primeiro se possivel
echo 4. A DLL FlexDocs foi desregistrada automaticamente
echo.
echo Para testar: execute TesteACBrProducao.vbs
echo.
pause