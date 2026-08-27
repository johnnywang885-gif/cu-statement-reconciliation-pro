@echo off
chcp 950 >nul
setlocal

where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    echo ================================================
    echo    ���~�G�z���t�Ω|���w�� PowerShell 7
    echo ================================================
    echo.
    echo ���u��ݭn�w�� PowerShell 7 ^(pwsh.exe^)�C
    echo �ЦܤU�C���}�U���w�ˡG
    echo https://github.com/PowerShell/PowerShell/releases
    echo.
    echo �w�˧�����Э��s���榹�{���C
    echo.
    pause
    exit /b
)


set "SELECTED_CUB="
set "DEFAULT_PWD=thifincub"
set "CUB_PWD=%DEFAULT_PWD%"

:SELECT_FILE
for /f "tokens=*" %%i in ('pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0select_cub.ps1"') do set "NEW_CUB=%%i"

if "%NEW_CUB%"=="" (
    if "%SELECTED_CUB%"=="" (
        echo.
        echo ������ɮסA�{�������C
        timeout /t 3 /nobreak >nul
        exit /b
    ) else (
        echo.
        echo �������ܡA�~��ϥ��ɮסG%SELECTED_CUB%
        timeout /t 2 /nobreak >nul
        goto MENU
    )
)
set "SELECTED_CUB=%NEW_CUB%"
set "CUB_PWD=%DEFAULT_PWD%"
goto GET_PASSWORD

:GET_PASSWORD
if not "%CUB_PWD%"=="" goto MENU
cls
echo.
echo ================================================
echo    CUB.MDB �K�X
echo ================================================
echo.
for /f "tokens=*" %%p in ('pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0get_password.ps1"') do set "CUB_PWD=%%p"
if "%CUB_PWD%"=="" (
    echo.
    echo �K�X���i�ťաA�Э��s��J�C
    timeout /t 2 /nobreak >nul
    goto GET_PASSWORD
)
goto MENU

:MENU
cls
echo.
echo ================================================
echo    ��b�� - �x������ƳB�z�t��
echo ================================================
echo.
echo   �ثe CUB.MDB�G%SELECTED_CUB%
echo.
echo   1. ���`����
echo   2. �ʤ���z��
echo   3. �@��z��e5%%�]�t���l�^
echo   4. ���� CUB.MDB
echo   5. ���͹�b�� PDF
echo   6. ����
echo.
echo ================================================
set /p CHOICE=�п�J�ﶵ (1-6):

if "%CHOICE%"=="1" goto RUN_ANOMALY
if "%CHOICE%"=="2" goto FILTER_PERCENT
if "%CHOICE%"=="3" goto RUN_TOP5
if "%CHOICE%"=="4" goto SELECT_FILE
if "%CHOICE%"=="5" goto GENERATE_PDF
if "%CHOICE%"=="6" goto END
goto MENU

:RUN_ANOMALY
cls
echo.
echo ���b���沧�`����...
echo.
pwsh -ExecutionPolicy Bypass -File "%~dp0find_anomaly_members.ps1" -CubPath "%SELECTED_CUB%" -CubPassword "%CUB_PWD%"
echo.
pause
goto MENU

:FILTER_PERCENT
cls
echo.
echo ���b�̤�ҿz��...
echo.
set /p PCT=�п�J�z��ʤ��� (1-100):
if "%PCT%"=="" goto FILTER_PERCENT
set /a PCT_NUM=%PCT% 2>nul
if %PCT_NUM% leq 0 goto FILTER_PERCENT
if %PCT_NUM% gtr 100 goto FILTER_PERCENT
pwsh -ExecutionPolicy Bypass -File "%~dp0filter_by_percent.ps1" -Percent %PCT%
echo.
pause
goto MENU

:RUN_TOP5
cls
echo.
echo [1/2] ���沧�`����...
echo.
pwsh -ExecutionPolicy Bypass -File "%~dp0find_anomaly_members.ps1" -CubPath "%SELECTED_CUB%" -CubPassword "%CUB_PWD%"
if %errorlevel% neq 0 (
    echo.
    echo ���`�������ѡA����C
    pause
    goto MENU
)
echo.
echo [2/2] �z��e5%%...
echo.
pwsh -ExecutionPolicy Bypass -File "%~dp0filter_top5.ps1" -Percent 5
echo.
pause
goto MENU

:GENERATE_PDF
cls
echo.
echo ���b���͹�b�� PDF...
echo.
if "%SELECTED_CUB%"=="" (
    echo ���~�G�|���ܤ@�� CUB.MDB�A�Э���ܡC
    pause
    goto SELECT_FILE
)
pwsh -ExecutionPolicy Bypass -File "%~dp0generate_statements.ps1" -CubPath "%SELECTED_CUB%" -CubPassword "%CUB_PWD%"
echo.
pause
goto MENU

:END
cls
echo.
echo �A���I
timeout /t 2 /nobreak >nul