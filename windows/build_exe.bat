@echo off
chcp 65001 >nul
echo ============================================
echo   CREAZIONE ESEGUIBILE - CONTA POLLINICA
echo ============================================
echo.

REM Verifica che Python sia installato
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato.
    echo Scarica Python da https://www.python.org/downloads/
    echo Assicurati di selezionare "Add Python to PATH" durante l'installazione.
    pause
    exit /b 1
)

echo [1/3] Installazione dipendenze...
pip install openpyxl pyinstaller sv-ttk python-docx
if errorlevel 1 (
    echo ERRORE durante l'installazione delle dipendenze.
    pause
    exit /b 1
)

echo.
echo [2/3] Creazione eseguibile...
cd /d "%~dp0"

REM Lettura vocale: inclusa solo se il modello vocale e' in codice\modelli
REM (vedi ISTRUZIONI_WINDOWS.txt). Senza modello l'exe funziona, ma senza voce.
set VOCE_OPTS=
if exist "..\codice\modelli\vosk-model*" (
    echo Modello vocale trovato: includo la lettura vocale.
    pip install vosk sounddevice pyttsx3
    if errorlevel 1 (
        echo ATTENZIONE: librerie vocali non installate, creo l'exe senza voce.
    ) else (
        set VOCE_OPTS=--add-data "..\codice\modelli;modelli" --collect-all vosk --collect-all sounddevice --collect-all _sounddevice_data --hidden-import pyttsx3.drivers --hidden-import pyttsx3.drivers.sapi5 --hidden-import comtypes.client --hidden-import win32com.client --hidden-import pythoncom
    )
) else (
    echo Modello vocale non trovato in codice\modelli: creo l'exe senza voce.
)

pyinstaller --onefile --windowed ^
  --add-data "..\codice\Polline_Template_Settimanale.xlsx;." ^
  --add-data "..\codice\concentrazioni_polliniche.xlsx;." ^
  --add-data "..\codice\ITA_Template_Bollettino_pubblicazione.docx;." ^
  --add-data "..\codice\ENG_Template_Bollettino_pubblicazione.docx;." ^
  --hidden-import dominio ^
  --hidden-import sessione ^
  --hidden-import esportatori ^
  --hidden-import percorsi ^
  --hidden-import voce ^
  --hidden-import docx ^
  --hidden-import sv_ttk ^
  %VOCE_OPTS% ^
  --name "Conta_Pollinica" ^
  ..\codice\polline_counter_gui.py
if errorlevel 1 (
    echo ERRORE durante la creazione dell'eseguibile.
    pause
    exit /b 1
)

echo.
echo [3/3] Pulizia file temporanei...
rmdir /s /q build 2>nul
del Conta_Pollinica.spec 2>nul
move dist\Conta_Pollinica.exe "%~dp0Conta_Pollinica.exe" 2>nul
rmdir /s /q dist 2>nul

echo.
echo ============================================
echo   FATTO!
echo ============================================
echo.
echo L'eseguibile si trova in:
echo   windows\Conta_Pollinica.exe
echo.
echo Per distribuirlo, copia SOLO il file .exe.
echo Il template Excel e' gia' incluso al suo interno.
echo.
pause
