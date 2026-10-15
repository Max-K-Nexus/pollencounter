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

REM Aspetto moderno (opzionale, solo estetico): se manca prova a installarlo in
REM silenzio; se non ci riesce il programma parte lo stesso con l'aspetto base.
python -c "import sv_ttk" >nul 2>&1
if errorlevel 1 (
    python -m pip install sv-ttk >nul 2>&1
)

REM Lettura vocale (opzionale): solo se il modello vocale e' stato messo in
REM codice\modelli. Se mancano le librerie prova a installarle; il programma
REM parte comunque anche senza voce.
if exist "%~dp0..\codice\modelli\vosk-model*" (
    python -c "import vosk, sounddevice, pyttsx3" >nul 2>&1
    if errorlevel 1 (
        echo Librerie della lettura vocale non trovate. Installazione in corso...
        python -m pip install vosk sounddevice pyttsx3
        if errorlevel 1 (
            echo ATTENZIONE: impossibile installare la lettura vocale.
            echo Il programma funziona, ma senza voce.
            echo Prova manualmente: python -m pip install vosk sounddevice pyttsx3
            pause
        )
    )
)

REM Avvia l'interfaccia grafica (pythonw = nessuna finestra console aggiuntiva)
start "" pythonw "%~dp0..\codice\polline_counter_gui.py"
