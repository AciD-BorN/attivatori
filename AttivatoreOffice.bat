@echo off
setlocal enabledelayedexpansion

goto check_Permissions

:check_Permissions
echo Eseguire questo programma come AMMINISTRATORE!!!. Controllo permessi in corso...

net session >nul 2>&1
if %errorLevel% == 0 (
echo Successo: Permessi Amministrativi confermati.
goto script
) else (
   echo Fallito: Permessi indeguati.
   goto promemoria
   )
pause >nul

echo .
echo .

:script
echo Attivazione Office by AciDBorN
echo ======================================
echo.

echo Ricerca dei files necessari in corso nel disco C:\ ...
echo Questo potrebbe richiedere alcuni minuti. Attendere prego...
echo.

dir /s /b "C:\ospp.vbs" > "%TEMP%\risultati_ricerca.txt" 2>nul

set trovati=0
set "percorso="

for /f "usebackq delims=" %%a in ("%TEMP%\risultati_ricerca.txt") do (
    set /a trovati+=1
    set "percorso=%%a"
)

if %trovati% EQU 0 (
    echo Nessun file trovato con il nome "ospp.vbs".
    echo Impossibile procedere con l'attivazione.
) else (
    echo Trovati files necessari!!!
    echo.
    echo Esecuzione del comando di attivazione...
    echo.
    
    cscript "!percorso!" /sethst:kms8.msguides.com
    cscript "!percorso!" /inpkey:BN8D3-W2QKT-M7Q73-Y3JWK-KQC63
    cscript "!percorso!" /act
    cscript "!percorso!" /dstatus
    
    echo.
    echo Attivazione eseguita. Controllare i risultati sopra.
)

echo.
@TIMEOUT /T 20 /NOBREAK
exit

:promemoria
echo.
echo.
echo Per eseguirlo come amministratore chiudere la finestra. Con il tasto destro del mouse
echo selezionare il file e lanciare con la voce "Esegui come Amministratore..."
echo.
echo.
@TIMEOUT /T 10 /NOBREAK
@exit