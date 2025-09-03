@echo off
title Rollback ACBrLibMDFe - AMBIENTE PRODUCAO
echo ===========================================
echo        ROLLBACK ACBrLibMDFe 
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
echo ATENCAO: Esta operacao ira DESFAZER a instalacao do ACBrLibMDFe
echo e restaurar o sistema anterior (FlexDocs se existir).
echo.
echo Pasta de trabalho: C:\Projetos\MDFe\NFE
echo.

REM Listar backups disponiveis
echo Backups disponiveis:
if exist "C:\Backup\MDFe\" (
    dir "C:\Backup\MDFe\" /b /ad
) else (
    echo Nenhum backup encontrado em C:\Backup\MDFe\
)

echo.
set /p dataBackup="Digite a data do backup para restaurar (formato: AAAA-MM-DD_HH-MM): "

if not exist "C:\Backup\MDFe\%dataBackup%" (
    echo.
    echo ERRO: Backup nao encontrado: C:\Backup\MDFe\%dataBackup%
    echo.
    echo Verifique os backups disponiveis acima e tente novamente.
    pause
    exit /b 1
)

echo.
set /p confirma="Confirma o ROLLBACK usando backup de %dataBackup%? (S/N): "
if /i "%confirma%" neq "S" (
    echo Rollback cancelado pelo usuario.
    pause
    exit /b 0
)

echo.
echo Executando rollback...

echo.
echo 1. Parando processos relacionados...
taskkill /f /im NFE.exe 2>nul
taskkill /f /im MDFE.exe 2>nul
timeout /t 3 /nobreak >nul

echo.
echo 2. Removendo arquivos ACBrLibMDFe...
del "ACBrMDFe32.dll" 2>nul
del "ACBrMDFe64.dll" 2>nul
del "ACBrLibMDFe.ini" 2>nul
del "libcrypto-1_1*.dll" 2>nul
del "libssl-1_1*.dll" 2>nul
del "libxml2.dll" 2>nul
del "libxslt.dll" 2>nul
del "libexslt.dll" 2>nul
del "libiconv.dll" 2>nul
del "TesteACBrProducao.vbs" 2>nul
rmdir "dep" /s /q 2>nul
echo ACBrLibMDFe removido.

echo.
echo 3. Restaurando arquivos do backup...
if exist "C:\Backup\MDFe\%dataBackup%\*.exe" (
    copy "C:\Backup\MDFe\%dataBackup%\*.exe" "." /Y
    echo Executaveis restaurados.
)

if exist "C:\Backup\MDFe\%dataBackup%\*.dll" (
    copy "C:\Backup\MDFe\%dataBackup%\*.dll" "." /Y
    echo DLLs restauradas.
)

if exist "C:\Backup\MDFe\%dataBackup%\*.ini" (
    copy "C:\Backup\MDFe\%dataBackup%\*.ini" "." /Y
    echo Configuracoes restauradas.
)

echo.
echo 4. Verificando se existe DLL FlexDocs para re-registrar...
if exist "MDFe_Util.dll" (
    echo Encontrada DLL FlexDocs, registrando...
    %windir%\Microsoft.NET\Framework\v4.0.30319\regasm "MDFe_Util.dll" /tlb:"MDFe_Util.tlb" /codebase
    if %errorlevel% equ 0 (
        echo ✓ DLL FlexDocs registrada com sucesso.
    ) else (
        echo ⚠ Aviso: Falha ao registrar DLL FlexDocs.
        echo   Pode ser necessario registro manual.
    )
) else (
    echo ✓ Nenhuma DLL FlexDocs para registrar.
)

echo.
echo 5. Testando sistema restaurado...
if exist "NFE.exe" (
    echo ✓ Executavel principal encontrado.
) else (
    echo ⚠ Executavel principal nao encontrado.
)

echo.
echo ===========================================
echo           ROLLBACK CONCLUIDO
echo ===========================================
echo.
echo Sistema restaurado para o estado anterior ao ACBrLibMDFe.
echo.
echo Arquivos restaurados de: C:\Backup\MDFe\%dataBackup%\
echo.
echo IMPORTANTE:
echo 1. Teste o funcionamento do sistema antes de usar
echo 2. Se havia certificado configurado, pode precisar reconfigurar
echo 3. Verifique se todas as funcionalidades estao operando
echo.
echo Se houver problemas, execute novamente o rollback ou
echo contacte o suporte tecnico.
echo.
pause