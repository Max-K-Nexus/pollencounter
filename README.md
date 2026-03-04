Pollencounter è uno strumento per la , pensato per supportare i monitoraggi aerobiologici e la produzione di bollettini.  
L’obiettivo è ridurre il lavoro manuale su file Excel, standardizzare i calcoli e rendere più semplice la gestione delle letture settimanali e dei report.

---



  parte del lavoro di compilazione dei file Excel per la conta pollinica.
  dovuti a copia/incolla e calcoli manuali.
  il formato dei dati e dei bollettini.
  sia l’uso “da tecnico” (script Python).

---



  che effettuano la conta pollinica settimanale.
  che lavorano con dati aerobiologici.
  che vogliono estendere o integrare lo strumento in altri flussi di lavoro.

Non è necessario essere programmatori per usare la versione .  
Per usare gli , è utile avere un minimo di familiarità con la riga di comando.

---



La struttura principale della cartella è:

Cartelle principali

    codice/  
    Contiene:

        i template Excel usati come base per i calcoli (Polline_Template_Settimanale.xlsx, concentrazioni_polliniche.xlsx)

        il file di configurazione pollencounter.cfg

        gli script Python principali:

            polline_counter.py → logica di elaborazione

            polline_counter_gui.py → versione con interfaccia grafica

    letture_settimanali/  
    Contiene i file Excel con le letture della conta pollinica per ogni settimana, ad esempio:

        Conta_Pollinica_16-02-2026.xlsx

        Conta_Pollinica_23-02-2026.xlsx

    riferimenti/  
    Contiene tabelle di monitoraggio storiche e di riferimento, ad esempio:

        Tabelle monitoraggi settimanali_2022_2023.xlsx

        Tabelle monitoraggi ultima settimana gennaio.xlsx

    script_aiuto/  
    Script di supporto (shell e Python) per automatizzare alcune operazioni, come:

        avvio rapido della conta

        applicazione di formattazioni ai file

    windows/  
    Tutto ciò che serve per l’uso su Windows:

        Conta_Pollinica.exe → eseguibile pronto all’uso

        AVVIA_CONTA_POLLINICA.bat → script per avviare l’applicazione

        ISTRUZIONI_WINDOWS.txt → istruzioni dettagliate per utenti Windows

        build_exe.bat → script per ricostruire l’eseguibile a partire dal codice (per sviluppatori)

File principali

    ISTRUZIONI.txt  
    Istruzioni generali per l’uso del progetto.

    CHANGELOG.md  
    Cronologia delle modifiche al progetto.

    esempio di bolletino.pdf  
    Esempio di bollettino prodotto a partire dai dati elaborati.

🔄 Flusso di lavoro tipico

Questa è l’idea generale di come si usa Pollencounter in pratica:

    Raccolta dati

        L’operatore compila i file Excel di conta settimanale nella cartella letture_settimanali/.

        I file seguono una convenzione di nome, ad esempio:
        Conta_Pollinica_DD-MM-YYYY.xlsx.

    Elaborazione

        L’utente avvia l’applicazione:

            tramite eseguibile Windows (Conta_Pollinica.exe)
            oppure

            tramite script Python (polline_counter.py o polline_counter_gui.py).

        L’applicazione legge:

            i file di lettura (letture_settimanali/)

            i template (codice/Polline_Template_Settimanale.xlsx, codice/concentrazioni_polliniche.xlsx)

            la configurazione (codice/pollencounter.cfg)

    Output

        Vengono generati file Excel aggiornati, tabelle e/o report pronti per essere usati nella redazione del bollettino.

        L’esempio di output è illustrato in esempio di bolletino.pdf.

🖥️ Uso per NON tecnici (Windows, eseguibile)

Questa è la modalità pensata per chi non vuole usare Python o la riga di comando.
Requisiti

    Sistema operativo: Windows

    Cartella del progetto completa (così come presente nella repository)

    Nessuna necessità di installare Python

Passi

    Scaricare il progetto

        Da GitHub, cliccare su Code → Download ZIP

        Estrarre lo ZIP in una cartella, ad esempio:
        C:\pollencounter

    Aprire la cartella windows/

        Percorso: C:\pollencounter\windows

    Leggere ISTRUZIONI_WINDOWS.txt

        Contiene le istruzioni specifiche per l’uso dell’eseguibile.

    Avviare l’applicazione

        Doppio clic su:

            Conta_Pollinica.exe  
            oppure

            AVVIA_CONTA_POLLINICA.bat

    Verificare i risultati

        Controllare i file generati/aggiornati nelle cartelle previste (ad esempio in letture_settimanali/ o in altre cartelle di output, se configurate).

🐍 Uso per utenti più tecnici (Python)

Questa modalità è pensata per chi vuole:

    eseguire direttamente gli script

    modificare il codice

    integrare Pollencounter in altri flussi di lavoro

Requisiti

    Python 3.x installato

    Alcune librerie Python (da documentare meglio, ad esempio: pandas, openpyxl, ecc.)

Passi

    Clonare o scaricare il repository
    bash

    git  https://github.com/Max-K-Nexus/pollencounter.git
     pollencounter

    Oppure scaricare lo ZIP da GitHub ed estrarlo.

    (Opzionale) Creare un ambiente virtuale
    bash

    python -m venv venv
     venv/bin/activate      
    .\venv\Scripts\activate       

    Installare le dipendenze  
    (Quando sarà disponibile un requirements.txt, ad esempio:)
    bash

    pip install -r requirements.txt

    Eseguire lo script principale

        Modalità riga di comando:
        bash

        python codice/polline_counter.py

        Modalità interfaccia grafica:
        bash

        python codice/polline_counter_gui.py

⚙️ File di configurazione pollencounter.cfg

Il file codice/pollencounter.cfg contiene le impostazioni principali del programma, ad esempio:

    percorsi delle cartelle di input/output

    parametri relativi ai template

    eventuali opzioni di calcolo

    Nota: la struttura esatta del file può essere documentata meglio in futuro, ma l’idea è che l’utente avanzato possa modificare questo file per adattare il comportamento del programma alle proprie esigenze (es. cambiare cartelle, nomi file, ecc.).

📁 Convenzioni sui nomi dei file

Per funzionare correttamente, il progetto si aspetta che i file seguano alcune convenzioni (esempio):

    File di lettura settimanale:

        Conta_Pollinica_DD-MM-YYYY.xlsx

    Template:

        Polline_Template_Settimanale.xlsx

        concentrazioni_polliniche.xlsx

    Suggerimento: mantenere i nomi originali dei file forniti nel progetto, a meno di sapere esattamente dove aggiornare i riferimenti nel codice o nella configurazione.

🧰 Script di aiuto

Nella cartella script_aiuto/ sono presenti alcuni script che possono:

    avviare rapidamente la conta pollinica (versione riga di comando o GUI)

    applicare formattazioni ai file Excel

    automatizzare piccoli passaggi ripetitivi

Esempi:

    AVVIA_CONTA_POLLINICA.sh

    AVVIA_CONTA_POLLINICA_GUI.sh

    applica_formattazione.py

Questi script sono pensati soprattutto per ambienti Linux/macOS o per utenti più esperti.
🧪 Esempio di output

Il file:

    esempio di bolletino.pdf

mostra un esempio di bollettino che può essere prodotto a partire dai dati elaborati con Pollencounter.
Può essere usato come riferimento per capire il tipo di risultato atteso.
🛠️ Per sviluppatori

Se vuoi modificare o estendere il progetto:

    Il codice principale è in codice/polline_counter.py e codice/polline_counter_gui.py.

    L’eseguibile Windows (Conta_Pollinica.exe) è probabilmente generato a partire da questi script (ad esempio con strumenti come pyinstaller), usando lo script windows/build_exe.bat.

    Le tabelle e i template Excel sono parte integrante della logica: modificare la struttura dei file può richiedere modifiche al codice.

❓ Domande frequenti (FAQ)

➤ Devo sapere programmare per usare Pollencounter?  
No, se usi la versione Windows eseguibile. Ti basta seguire le istruzioni in windows/ISTRUZIONI_WINDOWS.txt.

➤ Posso usare Pollencounter su macOS o Linux?  
Sì, ma in quel caso è consigliato usare gli script Python (polline_counter.py / polline_counter_gui.py) e non l’eseguibile Windows.

➤ Cosa succede se cambio i nomi dei file o delle cartelle?  
Il programma potrebbe non trovare più i file necessari. È consigliato mantenere la struttura e i nomi originali, oppure aggiornare di conseguenza il file di configurazione e/o il codice.

➤ Posso aggiungere nuove settimane di lettura?  
Sì. Basta aggiungere nuovi file Excel nella cartella letture_settimanali/, seguendo la stessa struttura e convenzione dei file esistenti.
🧯 Risoluzione problemi (troubleshooting)

    L’eseguibile non parte su Windows

        Verifica di aver estratto correttamente lo ZIP.

        Controlla che l’antivirus non blocchi l’eseguibile.

        Leggi windows/ISTRUZIONI_WINDOWS.txt per eventuali note specifiche.

    Gli script Python danno errore

        Controlla di avere Python 3.x installato.

        Verifica di aver installato tutte le librerie richieste.

        Assicurati di eseguire i comandi dalla cartella giusta (pollencounter/).

    Il programma non trova i file Excel

        Controlla che i file siano nelle cartelle corrette (letture_settimanali/, codice/, ecc.).

        Verifica che i nomi dei file non siano stati modificati.

🤝 Contributi

Al momento il progetto è pensato principalmente per uso interno / didattico, ma contributi e suggerimenti sono benvenuti.

Per proporre modifiche:

    Fai un fork della repository.

    Crea un branch con la tua modifica.

    Apri una pull request descrivendo:

        cosa hai cambiato

        perché

        come testare la modifica

📜 Licenza

(Da definire)

Se vuoi rendere il progetto liberamente utilizzabile, una scelta comune è la licenza MIT.
Puoi creare un file LICENSE con il testo della licenza MIT e aggiornare questa sezione di conseguenza.
