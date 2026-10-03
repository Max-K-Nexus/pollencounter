@echo off
chcp 65001 >nul

REM Verifica che Python sia installato
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato.
    echo Scarica Python da https://www.python.org/downloads/
    echo IMPORTANTE: durante l'installazione spunta "Add Python to PATH".
    pause
    exit /b 1
)

REM Verifica che openpyxl sia installato; se manca, lo installa automaticamente
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo openpyxl non trovato. Installazione in corso...
    python -m pip install openpyxl
    if errorlevel 1 (
        echo ERRORE: impossibile installare openpyxl.
        echo Prova manualmente: python -m pip install openpyxl
        pause
        exit /b 1
    )
)

REM Verifica python-docx (serve per i bollettini Word); se manca prova a
REM installarlo, ma il programma parte comunque anche senza
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo python-docx non trovato. Installazione in corso...
    python -m pip install python-docx
    if errorlevel 1 (
        echo ATTENZIONE: impossibile installare python-docx.
        echo Il programma funziona, ma non potra' generare i bollettini Word.
        echo Prova manualmente: python -m pip install python-docx
        pause
    )
)

REM Avvia l'interfaccia grafica (pythonw = nessuna finestra console aggiuntiva)
start "" pythonw "%~dp0..\codice\polline_counter_gui.py"
