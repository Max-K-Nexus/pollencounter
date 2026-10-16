#!/usr/bin/env python3
"""
Conta Pollinica — interfaccia a riga di comando.

Punto d'accesso alternativo sullo stesso dominio usato dalla GUI
(dominio.py, sessione.py, esportatori.py): non e' piu' il motore di cui
la GUI dipende (vedi CLAUDE.md / CHANGELOG per il contesto della riscrittura).
"""

import sys
from datetime import datetime, timedelta
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("ERRORE: openpyxl non installato. Installa con:")
    print("  pip3 install openpyxl")
    sys.exit(1)

import archivio
import dominio
import esportatori
import percorsi
import sessione
import voce

RADICE = percorsi.EXE_DIR


# ============================================================
# Menu e display
# ============================================================
def display_menu():
    print("\n" + "=" * 60)
    print("CONTA POLLINICA - SISTEMA AUTOMATIZZATO")
    print("=" * 60)
    print("\nCodici disponibili:\n")
    keys = list(dominio.CODICI_SPECIE.keys())
    for i in range(0, len(keys), 2):
        code1, specie1 = keys[i], dominio.CODICI_SPECIE[keys[i]]
        line = f"  {code1}: {specie1:<30}"
        if i + 1 < len(keys):
            code2, specie2 = keys[i + 1], dominio.CODICI_SPECIE[keys[i + 1]]
            line += f"  {code2}: {specie2}"
        print(line)
    print("\n" + "-" * 60)
    print("Comandi:")
    print("  01-59   Inserisce la specie corrispondente")
    print("  NNxQ    Inserisce Q occorrenze (es. 48x4 = 4 Alternaria)")
    print("  .       Ripete l'ultimo codice inserito")
    print("  r       Riepilogo giornata corrente")
    print("  w       Riepilogo settimanale (tutti i giorni)")
    print("  l       Ultimi inserimenti (storico)")
    print("  c       Correggi un giorno precedente")
    print("  n       Aggiungi una nota per la giornata")
    print("  f       Cambia il fattore di conversione (concentrazioni)")
    print("  u       Annulla ultimo inserimento")
    print("  b       Attiva/disattiva beep sonoro")
    print("  h       Mostra questo menu")
    print("  s       Salva il file (senza uscire)")
    print("  d       Chiudi giornata (puoi continuare con un altro giorno)")
    print("  q       Salva il file e esci")
    print("-" * 60 + "\n")


def _beep():
    """Emette un beep sonoro (voce.beep: winsound su Windows, tono WAV su
    Linux/macOS; bell ASCII come ultimo ripiego)."""
    if not voce.beep():
        print("\a", end="", flush=True)


# ============================================================
# Configurazione cartella di lavoro
# ============================================================
def carica_o_crea_config():
    global RADICE

    cartella = sessione.leggi_radice_archivio(percorsi.CONFIG_FILE)
    if cartella:
        print(f"  Cartella di lavoro: {cartella}")
        RADICE = cartella
        return

    print("\nConfigurazione della cartella di lavoro.")
    print("  Il programma crea al suo interno una cartella per ogni anno e,")
    print("  dentro quella, una per ogni settimana (es. 2026/S42_2026-10-12)")
    print("  con i dati e i bollettini di quella settimana.")
    print(f"  Invio = cartella corrente ({RADICE})")
    risposta = input("  Inserisci il percorso (o premi Invio per usare la cartella corrente): ").strip()

    if not risposta:
        cartella = RADICE
    else:
        cartella = Path(risposta)
        try:
            cartella.mkdir(parents=True, exist_ok=True)
        except Exception:
            print("  Percorso non valido, uso la cartella corrente.")
            cartella = RADICE

    sessione.salva_radice_archivio(percorsi.CONFIG_FILE, cartella)
    print(f"  [OK] Configurazione salvata in: {percorsi.CONFIG_FILE.name}")
    RADICE = cartella


# ============================================================
# Scelta file/sessione di partenza
# ============================================================
def chiedi_settimana(settimana_esistente=None):
    """Chiede l'intervallo settimanale. Ritorna la data del lunedi' (datetime)."""
    oggi = datetime.now()
    lun_corrente = dominio.lunedi_di(oggi)
    dom_corrente = lun_corrente + timedelta(days=6)

    if settimana_esistente:
        lun_file = dominio.lunedi_di(settimana_esistente)
        dom_file = lun_file + timedelta(days=6)
        print(f"\nIl file si riferisce alla settimana:")
        print(f"  dal {lun_file.strftime('%d-%m-%Y')} al {dom_file.strftime('%d-%m-%Y')}")
        risp = input("  Mantenere questa settimana? (s/n): ").strip().lower()
        if risp != "n":
            print(f"  -> Settimana: {lun_file.strftime('%d-%m-%Y')} (lun) "
                  f"- {dom_file.strftime('%d-%m-%Y')} (dom)")
            return lun_file

    default_str = f"dal {lun_corrente.strftime('%d-%m-%Y')} al {dom_corrente.strftime('%d-%m-%Y')}"
    print(f"\nChe settimana stiamo analizzando?")
    print(f"  Invio = settimana corrente ({default_str})")
    print(f"  Accetta: 9-2-2026  9/2/2026  09-02-26  9 feb 2026  ecc.")
    while True:
        inp = input("  Settimana: ").strip()
        if not inp:
            lunedi = lun_corrente
        else:
            dt = dominio.parse_data_flessibile(inp)
            if not dt:
                print("  Data non riconosciuta. Prova con: 9-2-2026 o 9/2/2026 o 9 feb 2026")
                continue
            lunedi = dominio.lunedi_di(dt)
        domenica = lunedi + timedelta(days=6)
        print(f"  -> Settimana: {lunedi.strftime('%d-%m-%Y')} (lun) "
              f"- {domenica.strftime('%d-%m-%Y')} (dom)")
        return lunedi


GIORNI_VALIDI = tuple(str(n) for n in dominio.GIORNI_NOMI)


def _stampa_giorni(lunedi):
    for num, nome in dominio.GIORNI_NOMI.items():
        data_giorno = lunedi + timedelta(days=num - 1)
        print(f"  {num}) {nome.upper():<12} {data_giorno.strftime('%d-%m-%Y')}")


def chiedi_giorno(lunedi):
    print("\nSeleziona il giorno di lavoro:")
    _stampa_giorni(lunedi)
    while True:
        scelta = input("Scegli (1-7): ").strip()
        if scelta in GIORNI_VALIDI:
            return int(scelta)
        print("Scelta non valida, riprova.")


def chiedi_ripresa_o_nuovo(radice):
    """Mostra il menu di ripresa: una riga per settimana, il lavoro interrotto e
    il file .xlsx della stessa settimana nella stessa voce (come nella GUI).
    Ritorna ('journal', path) / ('xlsx', path) / ('nuovo', None)."""
    info_journal = dict(sessione.recupera_sessioni(radice))
    fuori_schema = [f for f in sessione.cerca_file_ripresa(
        radice, esportatori.TEMPLATE_FILE.name) if f.parent == radice]
    voci = archivio.elenca(radice, legacy_xlsx=fuori_schema)

    if not voci:
        return "nuovo", None

    print("\nLavori riprendibili:")
    opzioni = []
    for voce in voci:
        etichetta = (f"settimana dal {voce.lunedi.strftime('%d-%m-%Y')}"
                     if voce.lunedi else voce.xlsx.name)
        if voce.journal:
            info = info_journal.get(voce.journal)
            operazioni = f"{info['n_operazioni']} operazioni non salvate" if info \
                else "file di sessione danneggiato"
            print(f"  {len(opzioni) + 1}) [interrotta] {etichetta}  ({operazioni})")
            opzioni.append(("journal", voce.journal))
        else:
            dettagli = sessione.conta_righe_log(voce.xlsx)
            fuori = "  [da riordinare]" if voce.fuori_struttura else ""
            print(f"  {len(opzioni) + 1}) {etichetta}  {dettagli}{fuori}")
            if voce.xlsx.name != etichetta:
                print(f"       {voce.xlsx.name}")
            opzioni.append(("xlsx", voce.xlsx))
    print(f"  n) Nuovo file (dal template)")
    print(f"  i) Importa da altra cartella")

    while True:
        scelta = input("Scelta: ").strip().lower()
        if scelta == "n":
            return "nuovo", None
        if scelta == "i":
            raw = input("  Percorso file .xlsx da importare (invio per annullare): ").strip()
            if not raw:
                continue
            p = Path(raw)
            if p.is_file() and p.suffix.lower() == ".xlsx":
                return "xlsx", p
            print("  Percorso non valido.")
            continue
        if scelta.isdigit():
            idx = int(scelta) - 1
            if 0 <= idx < len(opzioni):
                return opzioni[idx]
        print("Scelta non valida, riprova.")


# ============================================================
# Riepiloghi a schermo
# ============================================================
def mostra_riepilogo_giorno(settimana, giorno_num):
    nome_giorno = dominio.GIORNI_NOMI[giorno_num].upper()
    print(f"\n  Riepilogo {nome_giorno}:")
    righe = settimana.righe_specie(giorno_num)

    def _sezione(etichetta, codici_sezione):
        print(f"    {etichetta}:")
        totale = 0
        trovato = False
        for codice, specie, val in righe:
            if codice not in codici_sezione:
                continue
            print(f"      [{codice}] {specie}: {val}")
            totale += val
            trovato = True
        if not trovato:
            print("      (nessun dato)")
        print(f"      --- Totale {etichetta.lower()}: {totale}")
        return totale

    tp = _sezione("POLLINI", dominio.POLLINI_CODICI)
    ts = _sezione("SPORE", dominio.SPORE_CODICI)
    if tp + ts > 0:
        print(f"    === TOTALE GIORNO: {tp + ts}")
    print()


def mostra_riepilogo_settimana(settimana):
    intestazione = "  " + " " * 28 + "LUN  MAR  MER  GIO  VEN  SAB  DOM"
    print(f"\n{intestazione}")
    print("  " + "-" * 63)

    def _stampa_sezione(codici, etichetta_totale):
        has_data = False
        for codice in codici:
            vals = settimana.conteggi[codice]
            if any(v > 0 for v in vals):
                has_data = True
                sp = dominio.CODICI_SPECIE[codice][:22]
                vs = "".join(f"{v:5}" if v > 0 else "    -" for v in vals)
                print(f"  [{codice}] {sp:<24} {vs}")
        if not has_data:
            print("  (nessun dato)")
        ts = "".join(f"{v:5}" for v in dominio.somma_codici(settimana.conteggi, codici))
        print(f"  {etichetta_totale:<28} {ts}")

    _stampa_sezione(dominio.POLLINI_CODICI, "--- Totale pollini")
    print()
    _stampa_sezione(dominio.SPORE_CODICI, "--- Totale spore")
    print()


def mostra_storico(storico):
    if not storico:
        print("\n  (nessun inserimento in questa sessione)\n")
        return
    print("\n  Ultimi inserimenti:")
    for data, codice, _specie, quantita, ora in storico[-10:]:
        print(f"    {ora}  {dominio.descrivi_inserimento(codice, quantita)}  ({data})")
    print()


def correggi_giorno(settimana, journal):
    print("\nQuale giorno vuoi correggere?")
    _stampa_giorni(settimana.lunedi)

    scelta = input("Scegli (1-7, invio per annullare): ").strip()
    if scelta not in GIORNI_VALIDI:
        return
    giorno_num = int(scelta)
    nome_giorno = dominio.GIORNI_NOMI[giorno_num].upper()

    print(f"\n  Dati {nome_giorno} {settimana.data_str_giorno(giorno_num)}:")
    righe = settimana.righe_specie(giorno_num)
    if not righe:
        print("    (nessun dato)")
        return
    for codice, specie, val in righe:
        print(f"    [{codice}] {specie}: {val}")

    codice = dominio.normalizza_codice(input("\n  Codice da correggere (invio per annullare): ").strip())
    if not codice:
        return
    if not dominio.codice_valido(codice):
        print(f"  Codice non riconosciuto: {codice}")
        return

    val_attuale = settimana.conteggi[codice][giorno_num - 1]
    specie = dominio.CODICI_SPECIE[codice]
    nuovo = input(f"  {specie}: {val_attuale} -> nuovo valore: ").strip()
    if not nuovo:
        return
    try:
        nuovo_val = int(nuovo)
    except ValueError:
        print("  Valore non valido.")
        return
    try:
        settimana.correggi(giorno_num, codice, nuovo_val, journal=journal)
    except ValueError as e:
        print(f"  {e}")
        return
    print(f"  Corretto: [{codice}] {specie}: {val_attuale} -> {nuovo_val}")


def cambia_fattore(settimana, journal):
    print(f"\n  Fattore di conversione attuale: {settimana.fattore:g}")
    testo = input("  Nuovo fattore (invio per lasciarlo invariato): ").strip()
    if not testo:
        return
    try:
        nuovo = dominio.interpreta_fattore(testo)
    except ValueError as e:
        print(f"  {e}")
        return
    if abs(nuovo - settimana.fattore) < 1e-12:
        print("  Il fattore non e' cambiato.")
        return
    vecchio = settimana.imposta_fattore(nuovo, journal=journal)
    print(f"  Fattore: {vecchio:g} -> {nuovo:g} "
          f"(vale per il bollettino e per il file salvato di questa settimana)")


def aggiungi_nota(settimana, giorno_num, journal):
    nota = input("  Nota: ").strip()
    if not nota:
        print("  (nessuna nota inserita)")
        return
    settimana.aggiungi_nota(giorno_num, nota, journal=journal)
    print(f"  Nota registrata: {nota}")


# ============================================================
# Salvataggio
# ============================================================
def chiedi_percorso_salvataggio(settimana):
    """Percorso del .xlsx: per difetto quello dell'archivio (sempre lo stesso
    per quella settimana). Si puo' comunque indicarne un altro."""
    canonico = archivio.percorso_settimanale(RADICE, settimana.lunedi)
    try:
        mostrato = canonico.relative_to(RADICE)
    except ValueError:
        mostrato = canonico
    print(f"\n  File: {mostrato}")
    risposta = input("  Invio per salvare qui, oppure indica un altro percorso: ").strip()
    if not risposta:
        return canonico
    p = Path(risposta)
    if p.suffix.lower() != ".xlsx":
        p = p.with_suffix(".xlsx")
    if not p.is_absolute():
        p = RADICE / p
    return p


def menu_uscita_salvataggio(settimana, journal, percorso_salvato):
    """Chiede se salvare e dove. Ritorna il percorso del file salvato, o None
    se si esce senza salvare."""
    print("\n" + "=" * 60)
    while True:
        risp = input("\nSalvare e uscire? (s/n): ").strip().lower()
        if risp == "s":
            percorso = None
            if percorso_salvato:
                risp_rapido = input(f"  Salvare su '{percorso_salvato.name}'? (s = si, n = altro percorso): ").strip().lower()
                if risp_rapido == "s":
                    percorso = percorso_salvato
            if percorso is None:
                percorso = chiedi_percorso_salvataggio(settimana)
                canonico = archivio.percorso_settimanale(RADICE, settimana.lunedi)
                # Riscrivere il file della settimana nell'archivio e' il
                # comportamento normale: si chiede solo per gli altri percorsi.
                if percorso.exists() and percorso != canonico:
                    risp2 = input(f"  '{percorso.name}' esiste gia'. Sovrascrivere? (s/n): ").strip().lower()
                    if risp2 != "s":
                        continue

            esportatori.esporta_xlsx(settimana, percorso)
            print(f"\n  [OK] File salvato: {percorso}")
            journal.elimina()
            return percorso
        elif risp == "n":
            risp2 = input("Uscire senza salvare? I dati non salvati andranno persi. (s/n): ").strip().lower()
            if risp2 == "s":
                risp3 = input("  Conservare il lavoro per riprenderlo alla prossima sessione? (s/n): ").strip().lower()
                if risp3 == "s":
                    journal.chiudi()
                    print(f"  Il lavoro e' salvato in: {journal.path.name}")
                    print(f"  Puoi riprenderlo alla prossima esecuzione.")
                else:
                    journal.elimina()
                print("\n  [OK] Uscita senza salvataggio.")
                return None


def _scegli_duplicati_interattivo(data_str):
    print(f"\n  Il giorno '{data_str}' e' gia' presente nel riepilogo.")
    print("  Come gestire i duplicati?")
    print("    a) Sovrascrivere i dati esistenti")
    print("    b) Aggiungere una nuova riga")
    print("    c) Sommare ai dati esistenti")
    while True:
        risp = input("  Scelta (a/b/c): ").strip().lower()
        if risp in ("a", "b", "c"):
            return risp
        print("  Scelta non valida.")


# ============================================================
# Sessione di inserimento per un giorno
# ============================================================
def sessione_giorno(settimana, journal, giorno_num, stato):
    """Ritorna 'continue' (chiudi giornata) o 'quit' (uscire dal programma)."""
    nome_giorno = dominio.GIORNI_NOMI[giorno_num].upper()
    abbrev = dominio.GIORNI_ABBREV[giorno_num - 1]

    print(f"\n  Giorno:   {nome_giorno}")
    print(f"  Data:     {settimana.data_str_giorno(giorno_num)}")
    print()

    settimana.attiva_giorno(giorno_num, journal=journal)
    undo_consecutivi = 0

    while True:
        try:
            conteggio = settimana.totale_giorno(giorno_num)
            testo = input(f"{abbrev} [{conteggio}] >> ").strip()
        except KeyboardInterrupt:
            print(f"\n\n  Interrotto. {nome_giorno}: {settimana.totale_giorno(giorno_num)} osservazioni.")
            print(f"  Il lavoro resta salvato in: {journal.path.name}")
            return "quit"

        if not testo:
            continue

        cmd = dominio.interpreta_comando(testo)

        if isinstance(cmd, dominio.Azione):
            lettera = cmd.lettera
            if lettera == "q":
                print(f"\n  Chiusura {nome_giorno}: {settimana.totale_giorno(giorno_num)} osservazioni.")
                return "quit"
            if lettera == "d":
                print(f"\n  Chiusura {nome_giorno}: {settimana.totale_giorno(giorno_num)} osservazioni.")
                return "continue"
            if lettera == "h":
                display_menu()
            elif lettera == "r":
                mostra_riepilogo_giorno(settimana, giorno_num)
            elif lettera == "w":
                mostra_riepilogo_settimana(settimana)
            elif lettera == "l":
                mostra_storico(settimana.storico)
            elif lettera == "b":
                settimana.beep = not settimana.beep
                print(f"  Beep sonoro: {'ATTIVO' if settimana.beep else 'disattivo'}")
                if settimana.beep:
                    _beep()
            elif lettera == "c":
                correggi_giorno(settimana, journal)
            elif lettera == "n":
                aggiungi_nota(settimana, giorno_num, journal)
            elif lettera == "f":
                cambia_fattore(settimana, journal)
            elif lettera == "s":
                percorso = stato.get("percorso_salvato") or chiedi_percorso_salvataggio(settimana)
                esportatori.esporta_xlsx(settimana, percorso)
                stato["percorso_salvato"] = percorso
                print(f"  [OK] File salvato: {percorso}")
            elif lettera == "u":
                undo_consecutivi += 1
                if dominio.annullo_richiede_conferma(undo_consecutivi):
                    risp = input(f"  Hai annullato {undo_consecutivi - 1} inserimenti di fila. "
                                 f"Continuare? (s/n): ").strip().lower()
                    if risp != "s":
                        undo_consecutivi = 0
                        continue
                ris = settimana.annulla(giorno_num, journal=journal)
                if ris is None:
                    print("  Nessun inserimento da annullare.")
                else:
                    codice, qty, _ = ris
                    print(f"  <- Annullato: {dominio.descrivi_inserimento(codice, qty)}")
            continue

        undo_consecutivi = 0

        if isinstance(cmd, dominio.Ripeti):
            if not settimana.ultimo_codice:
                print("  Nessun codice precedente da ripetere.")
                continue
            codice, quantita = settimana.ultimo_codice, 1
        elif isinstance(cmd, dominio.Inserisci):
            codice, quantita = cmd.codice, cmd.quantita
        else:
            print(f"  {cmd.messaggio}" if cmd.messaggio else "  Comando non riconosciuto.")
            continue

        nuovo_val = settimana.inserisci(giorno_num, codice, quantita, journal=journal)
        print(f"  -> {dominio.descrivi_inserimento(codice, quantita)}  (totale giorno: {nuovo_val})")
        if settimana.beep:
            _beep()


# ============================================================
# Main
# ============================================================
def _esporta_protetto(cosa, esportazione):
    """Un errore qui non deve interrompere la chiusura: i conteggi sono gia'
    salvati nel .xlsx e l'operazione si puo' rifare riaprendo la settimana."""
    try:
        esportazione()
    except Exception as e:                      # noqa: BLE001 (qualunque errore)
        print(f"\n  ERRORE: non riesco a generare {cosa}: {e}")
        print("  Se il file e' aperto in Excel o in Word, chiudilo e riprova")
        print("  riaprendo la settimana: i conteggi sono salvati.")


def _esporta_annuale(settimana):
    """Riepilogo annuale: nella cartella dell'anno, uno solo per anno."""
    cartella = archivio.prepara_anno(RADICE, settimana.lunedi.year)
    percorso, n = esportatori.esporta_riepilogo_annuale(
        settimana, cartella, scegli_duplicati=_scegli_duplicati_interattivo,
        avvisa=lambda messaggio: print("\n  ATTENZIONE: " + messaggio.replace("\n", "\n  ")))
    if percorso:
        print(f"\n  [OK] Riepilogo annuale aggiornato: {percorso}")
        print(f"       {n} giorni esportati, fattore={settimana.fattore}")
    else:
        print("  Nessun dato da esportare nel riepilogo annuale.")


def _esporta_bollettini(settimana):
    """Bollettini ITA/ENG: nella cartella della settimana, accanto ai dati."""
    soglie = esportatori.carica_soglie(RADICE)
    if soglie is None:
        print(f"  ERRORE: file soglie '{esportatori.SOGLIE_FILE_NOME}' non trovato.")
        return
    cartella = archivio.prepara_settimana(RADICE, settimana.lunedi)
    creati = esportatori.genera_bollettini_word(settimana, cartella, soglie=soglie)
    if creati:
        for p in creati:
            print(f"  Bollettino: {p}")
        print("  Bollettini Word generati.")
    else:
        print("  Nessun dato presente: bollettini Word non generati.")


def _apri_journal(lunedi):
    """Journal nuovo per la settimana, in <radice>/in_corso/."""
    archivio.prepara_in_corso(RADICE)
    return sessione.Journal(archivio.percorso_journal(RADICE, lunedi))


def main():
    if not esportatori.TEMPLATE_FILE.exists():
        print(f"ERRORE: Template non trovato: {esportatori.TEMPLATE_FILE}")
        sys.exit(1)

    print("=" * 60)
    print("SISTEMA CONTA POLLINICA - RIEPILOGO SETTIMANALE")
    print("=" * 60)

    carica_o_crea_config()

    tipo, path = chiedi_ripresa_o_nuovo(RADICE)

    if tipo == "journal":
        settimana = sessione.ripristina_da_journal(path)
        journal = sessione.Journal.riprendi(path)
        print(f"\n  Sessione ripristinata: settimana del {settimana.lunedi.strftime('%d-%m-%Y')}")
        display_menu()
        lunedi = settimana.lunedi
    elif tipo == "xlsx":
        try:
            settimana = sessione.carica_da_xlsx(path)
        except ValueError as e:
            print(f"  ERRORE: {e}")
            sys.exit(1)
        print(f"\n  Ripreso: {path}")
        display_menu()
        lunedi = chiedi_settimana(settimana.lunedi)
        settimana.lunedi = lunedi
        journal = _apri_journal(lunedi)
        journal.avvia(settimana)
    else:
        display_menu()
        lunedi = chiedi_settimana()
        settimana = sessione.Settimana(lunedi)
        journal = _apri_journal(lunedi)
        journal.avvia(settimana)

    stato = {"percorso_salvato": None}

    while True:
        giorno_num = chiedi_giorno(lunedi)

        totale_esistente = settimana.totale_giorno(giorno_num)
        if totale_esistente > 0:
            nome_g = dominio.GIORNI_NOMI[giorno_num].upper()
            print(f"\n  ATTENZIONE: {nome_g} contiene gia' {totale_esistente} osservazioni.")
            risposta = input("  Continuare aggiungendo dati? (s/n): ").strip().lower()
            if risposta != "s":
                continue

        risultato = sessione_giorno(settimana, journal, giorno_num, stato)

        if risultato == "quit":
            break

        risposta = input("\nVuoi continuare con un altro giorno? (s/n): ").strip().lower()
        if risposta != "s":
            break

    percorso_finale = menu_uscita_salvataggio(settimana, journal, stato["percorso_salvato"])

    if percorso_finale:
        print("\n  Operazioni aggiuntive:")
        print("    1. Aggiorna riepilogo annuale")
        print("    2. Genera bollettini Word (ITA/ENG)")
        print("    3. Entrambi (annuale + bollettini)")
        print("    Invio = nessuna operazione aggiuntiva")
        risp_extra = input("\n  Scelta: ").strip()
        if risp_extra in ("1", "3"):
            _esporta_protetto("il riepilogo annuale", lambda: _esporta_annuale(settimana))
        if risp_extra in ("2", "3"):
            _esporta_protetto("i bollettini Word", lambda: _esporta_bollettini(settimana))

    print("\n  Sessione terminata.\n")

    if getattr(sys, "frozen", False):
        input("Premi INVIO per chiudere...")


if __name__ == "__main__":
    main()
