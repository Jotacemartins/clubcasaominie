@echo off
:: Cria tarefa agendada no Windows para rodar o rodar.py todo dia às 7h
:: Execute este arquivo UMA VEZ como Administrador

set PASTA=C:\Users\Helena CRM\Downloads\cezinha
set PYTHON=C:\Users\Helena CRM\AppData\Local\Microsoft\WindowsApps\python3.11.exe

schtasks /create /tn "ClubCasa - Automacao Diaria" ^
  /tr "%PYTHON% %PASTA%\rodar.py" ^
  /sc DAILY /st 07:00 ^
  /ru "Helena CRM" ^
  /f

echo.
echo Tarefa agendada com sucesso!
echo Vai rodar todo dia as 07:00 automaticamente.
echo.
pause
