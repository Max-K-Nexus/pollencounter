---
name: aggiorna-formattazione-template
description: Aggiorna la formattazione visiva (altezze righe, sfondi, bordi, grassetto) del template Excel del Conta Pollinica tramite applica_formattazione.py. Usare quando l'utente chiede di sistemare o aggiornare la formattazione/lo stile del template Excel.
---

# Utility di manutenzione template — `applica_formattazione.py`

Script standalone da eseguire **manualmente** quando si vuole aggiornare la
formattazione visiva del template Excel. Non è importato né chiamato dagli
script principali.

**Eseguire con:**
```bash
python3 script_aiuto/applica_formattazione.py
```

**Cosa fa:**
- Imposta altezze righe (15pt per righe 5–53 e 57–70)
- Applica sfondo verde tenue (`C5E0B4`) alle righe delle specie principali
- Applica bordi thin completi sulle sezioni dati
- Applica grassetto selettivo su nomi specie importanti e intestazioni
- Centra il contenuto delle celle dati
- Aggiorna le due copie del template (`codice/` e `windows/`)
- Crea automaticamente un backup prima di modificare il template principale

**Nota:** non inserisce formule Excel. Opera esclusivamente sulla formattazione.
