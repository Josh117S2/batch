@echo off
:: -----------------------------------------------------------------------------------------------------------------
:: Programa elaborado por LIC. JOSUE SANCHEZ SANCHEZ
:: TI SECTOR MADERO
:: Enero 2022
:: Programa que repara el problema del servicio de cola de impresión de Windows cuando este deja de funcionar. 
:: -----------------------------------------------------------------------------------------------------------------
setlocal ENABLEEXTENSIONS
MODE con:cols=80 lines=30
set spooldrvpath="C:\Windows\System32\spool\drivers\W32X86"
set spoolprntrspath="C:\Windows\System32\spool\PRINTERS"
set "_msg= Por favor espere..."
set "_psc=powershell -nop -ep bypass -c"
set "EchoRed=%_psc% write-host -back Black -fore Red"
set "EchoDRed=%_psc% write-host -back Black -fore DarkRed"
set "EchoGreen=%_psc% write-host -back Black -fore Green"
set "EchoGray=%_psc% write-host -back Black -fore Gray"
set "pd=C:\tempJS2"
if exist %pd% ( rd %pd% /S /Q )
::pause
%EchoDRed% ----------------------------------------------------------------------------
%EchoGreen% Iniciando Proceso de reparacion de spooler...
%EchoDRed% ----------------------------------------------------------------------------
ping -n 3 127.0.0.1 > nul
mkdir %pd%
reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Providers" > %pd%\Providers.txt
reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86\Drivers" > %pd%\Drivers.txt
reg query "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers" > %pd%\Printers.txt
:: Detengo el servicio de impresión
net stop spooler 2> nul
%EchoGreen% "Deteniendo servicio spooler."
ping -n 6 127.0.0.1 > nul
cls
%EchoDRed% ----------------------------------------------------------------------------
%EchoGreen% Spooler detenido.
:: Borrando contenido de la carpeta C:\Windows\System32\spool\drivers\W32X86
del "%spooldrvpath%\*" /F /Q /A 2> nul
for /F "eol=| delims=" %%I in ('dir "%spooldrvpath%\*" /AD /B 2^>nul') do rd /Q /S "%spooldrvpath%\%%I"
:: Borrando contenido de la carpeta C:\Windows\System32\spool\PRINTERS
del "%spoolprntrspath%\*" /F /Q /A 2> nul
for /F "eol=| delims=" %%I in ('dir "%spoolprntrspath%\*" /AD /B 2^>nul') do rd /Q /S "%spoolprntrspath%\%%I"
%EchoDRed% ----------------------------------------------------------------------------
%EchoGreen% "Borrando contenido de las carpetas: "
%EchoGray% %spooldrvpath%
%EchoGray% %spoolprntrspath%
%EchoDRed% ----------------------------------------------------------------------------
ping -n 3 127.0.0.1 > nul
:: Se realizo respaldo de la clave "Print" del registro del sistema
REG EXPORT HKLM\SYSTEM\CurrentControlSet\Control\Print "%pd%" 2> nul
%EchoGreen% "Se ha realizado el respaldo de la clave Print."
%EchoDRed% -----------------------------------------------------------------------------
ping -n 4 127.0.0.1 > nul
cls
:: Valido que el valor de Driver en "HKLM\SYSTEM\CurrentControlSet\Control\Print\Monitors\Local Port" es localspl.dll
:: Si no existe, la genero
%EchoDRed% ------------------------------------------------------------------------------
%EchoGreen% Configurando valores de registro de sistema.
%EchoDRed% ------------------------------------------------------------------------------
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Print\Monitors\Local Port" /v Driver 2^>^&1^|find "REG_"') do @set fn=%%b
If "%fn%" NEQ "localspl.dll" (reg delete "HKLM\SYSTEM\CurrentControlSet\Control\Print\Monitors\Local Port" /v Driver /f && reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Monitors\Local Port" /v Driver /t REG_SZ /d "localspl.dll" && %EchoGreen% Se cambio valor de Local Port -> localspl.dll ...  ok) 2> nul
%EchoGreen% Valor de Local Port -> localspl.dll ...  ok
%EchoDRed% ------------------------------------------------------------------------------
ping -n  127.0.0.1 > nul
::pause
%EchoGreen% "Configurando el registo del sistema clave Print ...  ok"
%EchoDRed% ------------------------------------------------------------------------------
:: Borro carpetas de la clave de registro HKEY_LOCAL_MACHINE\SYSTEM\CurrentContrlSet\Control\Print\
reg delete "HKLM\SYSTEM\ControlSet001\Control\Print\Environments\Windows NT x86\Print Processors" /f
reg delete "HKLM\SYSTEM\ControlSet001\Control\Print\Environments\Windows NT x86\Drivers\Version-3" /f
:: Seguido creo la clave de registro Providers con sus valores
reg delete "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /f
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers"
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /v EventLog /t REG_DWORD /d 00000003
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /v NetPopup /t REG_DWORD /d 00000001
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /v NetPopupToComputer /t REG_DWORD /d 00000000
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /v order /t REG_MULTI_SZ /d "LanMan Print Services"\0"Internet Print Provider"
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /v RestartJobOnPoolEnabled /t REG_DWORD /d 00000001
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /v RestartJobOnPoolError /t REG_DWORD /d 00000258
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers" /v RetryPopup /t REG_DWORD /d 00000000
:: Borro subcarpetas de la clave Drivers en HKLM\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86
reg delete "HKLM\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86\Drivers" /f
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86\Drivers"
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86\Drivers" /v Directory /d "W32X86"
:: Busco impresoras instaladas en esta ruta omitiendo las claves que se listan a continuación
for /f "tokens=*" %%A in (%pd%\Printers.txt) do (ECHO %%A |findstr /I /V "DefaultSpoolDirectory")>>%pd%\1.txt
for /f "tokens=*" %%A in (%pd%\1.txt) do (ECHO %%A |findstr /I /V "ResetDevmodesAttempts")>>%pd%\2.txt
for /f "tokens=*" %%A in (%pd%\2.txt) do (ECHO %%A |findstr /I /V "LANGIDOfLastDefaultDevmode")>>%pd%\3.txt
for /f "tokens=*" %%A in (%pd%\3.txt) do (ECHO %%A |findstr /I /V "Fax")>>%pd%\4.txt
for /f "tokens=*" %%A in (%pd%\4.txt) do (ECHO %%A |findstr /I /V "Microsoft")>>%pd%\5.txt
for /f "tokens=*" %%A in (%pd%\5.txt) do (ECHO %%A |findstr /I /V "OneNote")>>%pd%\6.txt
for /f "tokens=*" %%A in (%pd%\6.txt) do (ECHO %%A |findstr /I /V "pdf")>>%pd%\salida.txt 2> nul
more /E +1 %pd%\salida.txt > %pd%\printstodelete.txt
del /f /q "%pd%\salida.txt"
for /f "tokens=*" %%A in ( %pd%\printstodelete.txt ) do ( reg delete "%%A" /f )
::del /f /q "%pd%\1.txt" "%pd%\2.txt" "%pd%\3.txt" "%pd%\4.txt" "%pd%\5.txt" "%pd%\6.txt" > 2> nul
ping -n 5 127.0.0.1 > nul
cls
%EchoDRed% -----------------------------------------------------------------------------
%EchoGreen% Se eliminaron las claves de registro Providers, Windows NT x86
%EchoGreen% Drivers y Printers ...  ok
%EchoDRed% -----------------------------------------------------------------------------
ping -n 3 127.0.0.1 > nul
cls
:: Reiniciamos el sistema
%EchoDRed% -----------------------------------------------------------------------------
%EchoGreen% Presione cualquier tecla para reiniciar el sistema...
%EchoDRed% -----------------------------------------------------------------------------
pause
echo.
%EchoGreen% Borrando archivos temporales generados
rd %pd% /S /Q
::ping -n 5 127.0.0.1 > nul
%EchoDRed% -----------------------------------------------------------------------------
echo.
%EchoRed% Preparando para reiniciar, espere por favor...
echo.
%EchoDRed% -----------------------------------------------------------------------------

timeout 10 /nobreak
%EchoDRed% -----------------------------------------------------------------------------
ping -n 2 127.0.0.1 > nul
shutdown /r /f /t 000
%EchoDRed% -----------------------------------------------------------------------------
exit