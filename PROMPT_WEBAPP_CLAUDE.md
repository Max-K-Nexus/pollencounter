# Istruzioni per creare la web app "Conta Pollinica" con Claude

Copia il prompt qui sotto e incollalo direttamente nella chat di Claude (claude.ai)
chiedendo di creare un **artifact**. Claude generera' un singolo file HTML
autocontenuto che funziona in qualsiasi browser, senza installare nulla.

---

## Come usarlo

1. Vai su **claude.ai** (o usa l'API con artifacts abilitati)
2. Incolla il prompt nella chat
3. Claude resituira' un artifact: un file HTML interattivo
4. Scarica il file HTML e aprilo nel browser, oppure pubblicalo su qualsiasi
   hosting statico (GitHub Pages, Netlify, server universitario)

Per aggiornare l'app in futuro: riapri la chat con Claude, mostra il file HTML
corrente e chiedi le modifiche. Non serve ricompilare nulla.

---

## PROMPT DA INCOLLARE IN CLAUDE

---

Crea un artifact HTML autocontenuto (un unico file .html con CSS e JavaScript
integrati) per la **conta pollinica settimanale** usata in aerobiologia.

### Contesto d'uso

Gli operatori di laboratorio (non tecnici) osservano granuli di polline e spore
al microscopio e inseriscono un codice numerico per ogni granulo osservato.
Il software conta i granuli per specie e per giorno, calcola le concentrazioni
(granuli/m3) e genera un bollettino pollinico con livelli di allerta colorati.

---

### Flusso operativo

1. L'operatore sceglie la **settimana di riferimento** (seleziona la data del
   lunedi') e il **giorno corrente** (lun-dom).
2. Inserisce i codici specie uno alla volta in un campo di testo (campo grande,
   tasto Invio per confermare ogni codice). Ogni inserimento incrementa il
   contatore di quella specie per quel giorno.
3. Puo' correggere un errore con un tasto "Annulla ultimo".
4. Puo' passare a un altro giorno e continuare l'inserimento.
5. La tabella dei conteggi e il bollettino si aggiornano in tempo reale.
6. Al termine puo' stampare o scaricare il bollettino.

---

### Codici specie

I codici sono stringhe "01"-"59". Mostrare sempre il codice e il nome.
Pollini (01-47) e spore fungine (48-59):

```
01=ACERACEAE, 02=ALTRI POLLINI, 03=BETULACEAE, 04=Alnus, 05=Betula,
06=CANNABACEAE, 07=CHENO-AMAR, 08=COMPOSITAE, 09=Altre compositae,
10=Ambrosia, 11=Artemisia, 12=CORYLACEAE (somma c+o), 13=Carpinus/Ostrya,
14=Carpinus, 15=Ostrya carpinifolia, 16=Corylus avellana,
17=CUP-TAXACEAE, 18=ERICACEAE, 19=EUPHORBIACEAE, 20=FAGACEAE,
21=Castanea sativa, 22=Fagus sylvatica, 23=Quercus, 24=GRAMINEAE,
25=HIPPOCASTANACEAE, 26=JUGLANDACEAE, 27=LAURACEAE, 28=MIMOSACEAE,
29=MORACEAE, 30=MYRTACEAE, 31=OLEACEAE, 32=Altre oleaceae, 33=Fraxinus,
34=Ligustrum, 35=Olea, 36=PINACEAE, 37=PLANTAGINACEAE, 38=PLATANACEAE,
39=POLLINI NON IDENTIFICATI, 40=POLYGONACEAE, 41=SALICACEAE, 42=Populus,
43=Salix, 44=TILIACEAE, 45=ULMACEAE, 46=UMBELLIFERAE, 47=URTICACEAE,
48=Alternaria, 49=Botrytis, 50=Cladosporium, 51=Curvularia,
52=Epicoccum, 53=Helminthosporium, 54=Pithomyces, 55=Pleospora,
56=Polythrincium, 57=Stemphylium, 58=Tetraploa, 59=Torula
```

---

### Calcolo concentrazioni

- **Concentrazione giornaliera** = conteggio × fattore di conversione
- Il **fattore di conversione** e' un numero configurabile dall'utente
  (default: 0.4 p/m3 per granulo). Mostrarlo in un campo modificabile.
- **Media settimanale** = somma delle concentrazioni dei 7 giorni / 7
  (i giorni non ancora inseriti contano come 0)

---

### Soglie per il bollettino (valori p/m3)

Il bollettino mostra solo le famiglie che hanno almeno un codice inserito.
Per ogni famiglia, la media settimanale determina il livello:

| Famiglia | Assente (max) | Bassa (max) | Media (max) | Alta |
|---|---|---|---|---|
| Betulaceae | 0.5 | 15.9 | 49.9 | >50 |
| Composite | 0.0 | 4.9 | 24.9 | >25 |
| Corilacee | 0.5 | 15.9 | 49.9 | >50 |
| Fagaceae | 0.9 | 19.9 | 39.9 | >40 |
| Graminaceae | 0.5 | 9.9 | 29.9 | >30 |
| Oleaceae | 0.5 | 4.9 | 24.9 | >25 |
| Plantaginaceae | 0.0 | 0.4 | 1.9 | >2 |
| Urticaceae | 1.9 | 19.9 | 69.9 | >70 |
| Cupressaceae + Taxaceae | 3.9 | 29.9 | 89.9 | >90 |
| Cheno-Amarantaceae | 0.0 | 4.9 | 24.9 | >25 |
| Ulmaceae | 0.9 | 19.9 | 39.9 | >40 |
| Platanaceae | 0.9 | 19.9 | 39.9 | >40 |
| Aceracee | 0.9 | 19.9 | 39.9 | >40 |
| Pinaceae | 0.9 | 14.9 | 49.9 | >50 |
| Salicaceae | 0.9 | 19.9 | 39.9 | >40 |
| Alternaria (spora) | 1.0 | 19.0 | 100.0 | >100 |
| Cladosporium (spora) | 100.0 | 499.0 | 1000.0 | >1000 |

Mapping codice → famiglia per il bollettino:
```
01=Aceracee, 03=Betulaceae, 04=Betulaceae, 05=Betulaceae,
07=Cheno-Amarantaceae, 08=Composite, 09=Composite, 10=Composite,
11=Composite, 12=Corilacee, 13=Corilacee, 14=Corilacee, 15=Corilacee,
16=Corilacee, 17=Cupressaceae + Taxaceae, 20=Fagaceae, 21=Fagaceae,
22=Fagaceae, 23=Fagaceae, 24=Graminaceae, 31=Oleaceae, 32=Oleaceae,
33=Oleaceae, 34=Oleaceae, 35=Oleaceae, 36=Pinaceae, 37=Plantaginaceae,
38=Platanaceae, 41=Salicaceae, 42=Salicaceae, 43=Salicaceae,
45=Ulmaceae, 47=Urticaceae, 48=Alternaria, 50=Cladosporium
```

I codici non presenti in questo mapping (es. 02, 06, 25, 26...) vengono
conteggiati nella tabella ma non compaiono nel bollettino.

---

### Colori del bollettino

- **Assente:** verde (#00B050), testo bianco
- **Bassa:** giallo (#FFD966), testo nero
- **Media:** arancione (#F4B084), testo nero
- **Alta:** rosso (#FF0000), testo bianco

---

### Interfaccia richiesta

L'interfaccia deve avere **tre sezioni principali**, visibili contemporaneamente
o come schede (tab):

#### 1. Inserimento dati

- Campo di input grande e prominente per il codice (2 cifre, autofocus).
  Premendo Invio il codice viene registrato, il campo si svuota e rimane
  pronto per il prossimo inserimento.
- Selezione del giorno corrente (lun-dom) con la data corrispondente.
- Contatore "Inserimenti questa sessione: N".
- Pulsante "Annulla ultimo" per cancellare l'ultimo inserimento.
- Feedback visivo: quando si inserisce un codice valido, mostrare
  brevemente il nome della specie (0.5 secondi).
- Errore visivo se il codice non esiste (non 01-59).

#### 2. Tabella riassuntiva settimanale

- Righe: tutte le specie che hanno almeno un conteggio > 0.
- Colonne: specie | lun | mar | mer | gio | ven | sab | dom | totale settimanale.
- Mostrare sia il conteggio grezzo che la concentrazione (p/m3) per ogni cella.
- Le spore fungine in una sezione separata sotto i pollini.

#### 3. Bollettino pollinico

- Titolo: "BOLLETTINO POLLINICO - [mese] [anno]"
- Una riga per ogni famiglia/specie rilevata (quelle con conteggi > 0
  e presenti nel mapping soglie).
- Colonne: famiglia | lun | mar | mer | gio | ven | sab | dom | media (p/m3)
- Le celle con le concentrazioni giornaliere hanno il colore del livello.
- La cella "media" ha il colore del livello medio settimanale, con il
  valore numerico e l'etichetta (es. "3.2 — Bassa").
- Legenda colori in fondo.

---

### Persistenza dati

Salvare automaticamente i dati inseriti nel `localStorage` del browser, in modo
che ricaricare la pagina non perda i dati. Aggiungere un pulsante
"Nuova settimana" che svuota i dati dopo conferma.

---

### Export

- Pulsante "Stampa bollettino" che apre la finestra di stampa del browser
  con solo il bollettino formattato (senza i controlli dell'interfaccia).
- Pulsante "Esporta CSV" che scarica un file .csv con la tabella riassuntiva
  (utile per importare in Excel).

---

### Requisiti tecnici

- Singolo file HTML autocontenuto (CSS e JS inline, nessuna dipendenza esterna).
- Funziona offline dopo il primo caricamento.
- Responsive: utilizzabile sia su PC che su tablet (gli operatori usano
  spesso tablet in laboratorio).
- Testo dell'interfaccia interamente in italiano.
- Nessun framework esterno (no React, no Vue, no jQuery): solo HTML+CSS+JS vanilla.

---

## Note per estensioni future

Dopo aver ottenuto l'artifact base, si puo' chiedere a Claude ulteriori
miglioramenti in chat separata mostrando il file HTML corrente:

- **Export Excel (.xlsx):** aggiungere la libreria SheetJS (CDN) per generare
  un file Excel formattato simile al template originale.
- **Multi-sessione:** permettere di caricare dati di settimane precedenti
  da un file JSON esportato.
- **Backend Flask:** se si vuole un server centralizzato, chiedere a Claude
  di separare la logica in un file `app.py` Flask che serve la stessa
  interfaccia HTML e salva i dati su file o database SQLite.
- **Autenticazione:** aggiungere login semplice (username + password) se
  piu' laboratori condividono lo stesso server.
- **QR Code:** generare un QR code che punta alla pagina web, stampabile
  e affissibile in laboratorio.
