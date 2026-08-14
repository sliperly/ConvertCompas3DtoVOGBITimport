@echo off
chcp 65001 > nul
echo ========================================
echo  ДИАГНОСТИКА
echo ========================================
echo.

echo [1] Версия PowerShell:
powershell.exe -NoProfile -Command "Write-Host $PSVersionTable.PSVersion"

echo.
echo [2] Политика выполнения:
powershell.exe -NoProfile -Command "Write-Host (Get-ExecutionPolicy)"

echo.
echo [3] Файлы в папке скрипта:
dir /B "%~dp0"

echo.
echo [4] Тест Excel COM:
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "try { $xl = New-Object -ComObject Excel.Application; Write-Host 'Excel COM: OK'; $xl.Quit() } catch { Write-Host 'Excel COM ОШИБКА:' $_.Exception.Message }"

echo.
echo [5] Поиск файлов по маскам:
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$d='%~dp0'.TrimEnd('\'); @('*Состав*','*файлами*','*технолог*','*номенклатур*') | foreach { $f=Get-ChildItem $d -File | where {$_.Name -like $_}; if($f){Write-Host 'OK:' $f[0].Name}else{Write-Host 'НЕ НАЙДЕН:' $_} }"

echo.
echo [6] Полный запуск convert.ps1:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0convert.ps1"

echo.
pause
