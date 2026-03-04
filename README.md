# 🌿 Pollencounter

**Pollencounter** è uno strumento per la conta pollinica, pensato per supportare i monitoraggi aerobiologici e la produzione di bollettini ufficiali.

L’obiettivo principale è ridurre il lavoro manuale su file Excel, standardizzare i calcoli e rendere più semplice la gestione delle letture settimanali e dei report. Non è necessario essere programmatori per usare la versione Windows.

---

## ✨ Caratteristiche principali
* ✅ **Automazione**: Riduce drasticamente l'inserimento manuale e i calcoli ripetitivi.
* ✅ **Standardizzazione**: Calcoli uniformi per garantire la qualità dei dati aerobiologici.
* ✅ **Versatilità**: Utilizzabile tramite interfaccia grafica (GUI) o script Python.
* ✅ **Pronto all'uso**: Versione eseguibile per Windows inclusa, senza necessità di installare Python.

---

## 📂 Struttura della Repository

```text
pollencounter/
├── codice/                 # Core del software e configurazione
│   ├── pollencounter.cfg   # Impostazioni dei percorsi e parametri
│   ├── polline_counter.py  # Logica di elaborazione (riga di comando)
│   ├── polline_counter_gui.py # Versione con interfaccia grafica
│   ├── Polline_Template_Settimanale.xlsx # Template base per i calcoli
│   └── concentrazioni_polliniche.xlsx     # Template per le concentrazioni
├── letture_settimanali/    # Cartella di INPUT (file di conta settimanale)
│   ├── Conta_Pollinica_16-02-2026.xlsx
│   └── Conta_Pollinica_23-02-2026.xlsx
├── riferimenti/            # Tabelle storiche e dati di riferimento
├── script_aiuto/           # Utility per formattazione e avvio rapido (Bash/Python)
├── windows/                # Risorse specifiche per utenti Windows
│   ├── Conta_Pollinica.exe # Eseguibile pronto all'uso
│   ├── AVVIA_CONTA_POLLINICA.bat # Script di avvio rapido
│   ├── ISTRUZIONI_WINDOWS.txt     # Guida specifica per Windows
│   └── build_exe.bat       # Script per compilare l'eseguibile (dev)
├── esempio di bollettino.pdf # Esempio del risultato finale
├── CHANGELOG.md            # Cronologia delle modifiche
└── ISTRUZIONI.txt          # Documentazione generale
```

## 🔄 Flusso di Lavoro Tipico


1.  **Raccolta dati**: L’operatore compila i file Excel di conta settimanale nella cartella `letture_settimanali/` seguendo la convenzione di nome `Conta_Pollinica_GG-MM-AAAA.xlsx`.
2.  **Elaborazione**: L’utente avvia l’applicazione tramite l'eseguibile Windows (`Conta_Pollinica.exe`) o tramite script Python. Il programma legge i file di lettura, i template e la configurazione nella cartella `codice/`.
3.  **Output**: Vengono generati file Excel aggiornati e report pronti per la redazione del bollettino (come mostrato in `esempio di bollettino.pdf`).

---

## 🚀 Guida all'Uso

### 🪟 Utenti Windows (Non tecnici)
*Questa modalità non richiede l'installazione di Python.*

1.  **Scarica** il progetto da GitHub (`Code` -> `Download ZIP`) ed estrailo.
2.  Apri la cartella **`windows/`**.
3.  Leggi il file **`ISTRUZIONI_WINDOWS.txt`**.
4.  Avvia l'applicazione con un doppio clic su **`Conta_Pollinica.exe`** o **`AVVIA_CONTA_POLLINICA.bat`**.

### 🐍 Utenti Python (Sviluppatori)
*Per chi desidera modificare il codice o integrare lo script in altri flussi.*

1.  **Clona il repository**:
    ```bash
    git clone [https://github.com/Max-K-Nexus/pollencounter.git](https://github.com/Max-K-Nexus/pollencounter.git)
    cd pollencounter
    ```
2.  **Installa le dipendenze** (assicurati di avere `pandas` e `openpyxl` installati):
    ```bash
    pip install pandas openpyxl
    ```
3.  **Esegui lo script**:
    ```bash
    python codice/polline_counter_gui.py
    ```

---

## ⚙️ Configurazione e Convenzioni

* **File `.cfg`**: Il file `codice/pollencounter.cfg` permette di modificare i percorsi delle cartelle e i parametri di calcolo senza dover mettere mano al codice sorgente.
* **Nomenclatura**: È fondamentale mantenere i nomi dei file originali e la struttura delle cartelle. I file in `letture_settimanali/` devono tassativamente seguire il formato data indicato (`GG-MM-AAAA`).

---

## 🛠️ Risoluzione Problemi (FAQ)

* **L'eseguibile non parte?** Verifica di aver estratto correttamente l'archivio ZIP e che l'antivirus non stia bloccando il file eseguibile.
* **Errore nei file Excel?** Assicurati di non aver modificato o spostato la struttura delle colonne nei template presenti nella cartella `codice/`.
* **Uso su Mac/Linux?** L'eseguibile `.exe` è solo per Windows. Su altri sistemi è necessario utilizzare gli script Python presenti nella cartella `codice/`.

---

## 👥 Autori e Contributi

Il progetto **Pollencounter** è stato sviluppato da:

* **Simone Bettella** — Ideazione, sviluppo originale e logica di calcolo.
* **Massimiliano Iotti** — Manutenzione, automazione, documentazione e supporto Windows.

---

## 📜 Licenza

Il progetto è distribuito sotto licenza **GNU General Public License v3.0 (GPL‑3.0)**.
Se utilizzi Pollencounter in un tuo progetto o report, ti chiediamo gentilmente di citare gli autori:

> **Pollencounter** – sviluppato da Simone Bettella e Massimiliano Iotti  
> [https://github.com/Max-K-Nexus/pollencounter](https://github.com/Max-K-Nexus/pollencounter)
