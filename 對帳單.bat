@echo off
chcp 950 >nul
setlocal

:MENU
cls
echo.
echo ================================================
echo    對帳單工具 - 異常社員過濾與排序系統
echo ================================================
echo.
echo   1. 執行異常偵測
echo   2. 依比例過濾
echo   3. 不寄發對帳單管理
echo   4. 對帳單回覆登錄
echo   5. 查核登記管理
echo   6. 科目年度結餘
echo   7. 貸款查核合併
echo   8. 結束
echo.
echo ================================================
set /p CHOICE=請輸入選項 (1-8):

if "%CHOICE%"=="1" goto RUN_ANOMALY
if "%CHOICE%"=="2" goto FILTER_PERCENT
if "%CHOICE%"=="3" goto MANAGE_NONMAIL
if "%CHOICE%"=="4" goto RECONCILE_REPLY
if "%CHOICE%"=="5" goto AUDIT_REGISTER
if "%CHOICE%"=="6" goto BALANCE_CHECK
if "%CHOICE%"=="7" goto LOAN_MERGE
if "%CHOICE%"=="8" goto END
goto MENU

:RUN_ANOMALY
cls
echo.
echo 正在執行異常偵測...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0find_anomaly_members.ps1" %*
echo.
pause
goto MENU

:FILTER_PERCENT
cls
echo.
echo 正在依條件篩選 M_對帳明細...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0filter_by_percent.ps1" %*
echo.
pause
goto MENU

:MANAGE_NONMAIL
cls
echo.
echo 不寄發對帳單管理...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0manage_nonmail.ps1" %*
echo.
pause
goto MENU

:RECONCILE_REPLY
cls
echo.
echo 對帳單回覆登錄...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0reconcile_reply.ps1" %*
echo.
pause
goto MENU

:AUDIT_REGISTER
cls
echo.
echo 查核登記管理...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0audit_register.ps1" %*
echo.
pause
goto MENU

:BALANCE_CHECK
cls
echo.
echo 科目年度結餘...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0balance_check.ps1" %*
echo.
pause
goto MENU

:LOAN_MERGE
cls
echo.
echo 貸款查核合併作業...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0loan_merge_check.ps1" %*
echo.
pause
goto MENU

:END
cls
echo.
echo 再見！
timeout /t 2 /nobreak >nul
