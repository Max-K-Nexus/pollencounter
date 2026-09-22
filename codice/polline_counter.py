#!/usr/bin/env python3
"""
Conta Pollinica — interfaccia a riga di comando.

Punto d'accesso alternativo sullo stesso dominio usato dalla GUI
(dominio.py, sessione.py, esportatori.py): non e' piu' il motore di cui
la GUI dipende (vedi CLAUDE.md / CHANGELOG per il contesto della riscrittura).
"""

import sys
from datetime import timedelta
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("ERRORE: openpyxl non installato. Installa con:")
    print("  pip3 install openpyxl")
    sys.exit(1)

try:
    import winsound
except ImportError:
    winsound = None

import dominio
import esportatori
import percorsi
import sessione

OUTPUT_DIR = percorsi.EXE_DIR


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
    print("  u       Annulla ultimo inserimento")
    print("  b       Attiva/disattiva beep sonoro")
    print("  h       Mostra questo menu")
    print("  s       Salva il file (senza uscire)")
    print("  d       Chiudi giornata (puoi continuare con un altro giorno)")
    print("  q       Salva il file e esci")
    print("-" * 60 + "\n")


def _beep():
    """Emette un beep sonoro (winsound su Windows, bell ASCII come fallback)."""
    if winsound is not None:
        winsound.Beep(800, 150)
    else:
        print("\a", end="", flush=True)


# ============================================================
# Configurazione cartella di lavoro
# ============================================================
def carica_o_crea_config():
    global OUTPUT_DIR
    import datetime as _dt
    anno = _dt.datetime.now().year

    cartella = sessione.leggi_cartella_anno(percorsi.CONFIG_FILE, anno)
    if cartella:
        print(f"  Cartella di lavoro {anno}: {cartella}")
        OUTPUT_DIR = cartella
        return

    print(f"\nConfigurazione cartella di lavoro per l'anno {anno}.")
    print(f"  Tutti i file di questa stagione (settimanali e riepilogo annuale)")
    print(f"  verranno salvati in questa cartella.")
    print(f"  Invio = cartella corrente ({OUTPUT_DIR})")
    risposta = input("  Inserisci il percorso (o premi Invio per usare la cartella corrente): ").strip()

    if not risposta:
        cartella = OUTPUT_DIR
    else:
        cartella = Path(risposta)
        try:
            cartella.mkdir(parents=True, exist_ok=True)
        except Exception:
            print("  Percorso non valido, uso la cartella corrente.")
            cartella = OUTPUT_DIR

    sessione.salva_cartella_anno(percorsi.CONFIG_FILE, anno, cartella)
    print(f"  [OK] Configurazione salvata in: {percorsi.CONFIG_FILE.name}")
    OUTPUT_DIR = cartella


# ============================================================
# Scelta file/sessione di partenza
# ============================================================
def chiedi_settimana(settimana_esistente=None):
    """Chiede l'intervallo settimanale. Ritorna la data del lunedi' (datetime)."""
    import datetime as _dt
    oggi = _dt.datetime.now()
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


def chiedi_giorno(lunedi):
    print("\nSeleziona il giorno di lavoro:")
    for num, nome in dominio.GIORNI_NOMI.items():
        data_giorno = lunedi + timedelta(days=num - 1)
        print(f"  {num}) {nome.upper():<12} {data_giorno.strftime('%d-%m-%Y')}")
    while True:
        scelta = input("Scegli (1-7): ").strip()
        if scelta in [str(x) for x in range(1, 8)]:
            return int(scelta)
        print("Scelta non valida, riprova.")


def chiedi_ripresa_o_nuovo(cartella):
    """Mostra il menu di ripresa (sessioni journal + file .xlsx).
    Ritorna ('journal', path) / ('xlsx', path) / ('nuovo', None)."""
    sessioni = sessione.recupera_sessioni(cartella)
    file_esistenti = sessione.cerca_file_ripresa(cartella, esportatori.TEMPLATE_FILE.name)

    if not sessioni and not file_esistenti:
        return "nuovo", None

    print("\nSessioni recuperabili:")
    opzioni = []
    for path, info in sessioni:
        print(f"  {len(opzioni) + 1}) [interrotta] settimana del {info['lunedi']}  "
              f"({info['n_operazioni']} operazioni non salvate)")
        opzioni.append(("journal", path))
    for f in file_esistenti:
        print(f"  {len(opzioni) + 1}) {f.name}  {sessione.conta_righe_log(f)}")
        opzioni.append(("xlsx", f))
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
        totali = [0] * 7
        has_data = False
        for codice in codici:
            vals = settimana.conteggi[codice]
            if any(v > 0 for v in vals):
                has_data = True
                sp = dominio.CODICI_SPECIE[codice][:22]
                vs = "".join(f"{v:5}" if v > 0 else "    -" for v in vals)
                print(f"  [{codice}] {sp:<24} {vs}")
                for i in range(7):
                    totali[i] += vals[i]
        if not has_data:
            print("  (nessun dato)")
        ts = "".join(f"{v:5}" for v in totali)
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
    for data, codice, specie, quantita, ora in storico[-10:]:
        if quantita > 1:
            print(f"    {ora}  [{codice}] {specie} x{quantita}  ({data})")
        else:
            print(f"    {ora}  [{codice}] {specie}  ({data})")
    print()


def correggi_giorno(settimana, journal):
    print("\nQuale giorno vuoi correggere?")
    for num, nome in dominio.GIORNI_NOMI.items():
        data_giorno = settimana.data_giorno(num)
        print(f"  {num}) {nome.upper():<12} {data_giorno.strftime('%d-%m-%Y')}")

    scelta = input("Scegli (1-7, invio per annullare): ").strip()
    if not scelta or scelta not in [str(x) for x in range(1, 8)]:
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
def chiedi_percorso_salvataggio(nome_default):
    print(f"\n  File: {nome_default}")
    risposta = input("  Percorso (invio = nome predefinito nella cartella corrente): ").strip()
    if not risposta:
        return OUTPUT_DIR / nome_default
    p = Path(risposta)
    if p.suffix.lower() != ".xlsx":
        p = p.with_suffix(".xlsx")
    if not p.is_absolute():
        p = OUTPUT_DIR / p
    return p


def _nome_default(settimana, nome_ripreso):
    return nome_ripreso or f"Conta_Pollinica_{settimana.lunedi.strftime('%d-%m-%Y')}.xlsx"


def menu_uscita_salvataggio(settimana, journal, nome_ripreso, percorso_salvato):
    """Chiede se salvare e dove. Ritorna (True, cartella, percorso) o (False, None, None)."""
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
                nome_default = _nome_default(settimana, nome_ripreso)
                percorso = chiedi_percorso_salvataggio(nome_default)
                if percorso.exists():
                    risp2 = input(f"  '{percorso.name}' esiste gia'. Sovrascrivere? (s/n): ").strip().lower()
                    if risp2 != "s":
                        continue

            esportatori.esporta_xlsx(settimana, percorso)
            print(f"\n  [OK] File salvato: {percorso}")
            journal.elimina()
            return True, percorso.parent, percorso
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
                return False, None, None


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
    abbrev = dominio.GIORNI_NOMI[giorno_num][:3].upper()

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
                stato["uscita_richiesta"] = True
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
            elif lettera == "s":
                nome_default = _nome_default(settimana, stato["nome_ripreso"])
                percorso = stato.get("percorso_salvato") or chiedi_percorso_salvataggio(nome_default)
                esportatori.esporta_xlsx(settimana, percorso)
                stato["percorso_salvato"] = percorso
                print(f"  [OK] File salvato: {percorso}")
            elif lettera == "u":
                undo_consecutivi += 1
                if undo_consecutivi > 5 and (undo_consecutivi - 1) % 5 == 0:
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
                    specie = dominio.CODICI_SPECIE[codice]
                    if qty > 1:
                        print(f"  <- Annullato: [{codice}] {specie} x{qty}")
                    else:
                        print(f"  <- Annullato: [{codice}] {specie}")
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

        specie = dominio.CODICI_SPECIE[codice]
        nuovo_val = settimana.inserisci(giorno_num, codice, quantita, journal=journal)
        if quantita > 1:
            print(f"  -> [{codice}] {specie} x{quantita}  (totale giorno: {nuovo_val})")
        else:
            print(f"  -> [{codice}] {specie}  (totale giorno: {nuovo_val})")
        if settimana.beep:
            _beep()


# ============================================================
# Main
# ============================================================
def main():
    if not esportatori.TEMPLATE_FILE.exists():
        print(f"ERRORE: Template non trovato: {esportatori.TEMPLATE_FILE}")
        sys.exit(1)

    print("=" * 60)
    print("SISTEMA CONTA POLLINICA - RIEPILOGO SETTIMANALE")
    print("=" * 60)

    carica_o_crea_config()

    tipo, path = chiedi_ripresa_o_nuovo(OUTPUT_DIR)

    nome_ripreso = None
    if tipo == "journal":
        settimana = sessione.ripristina_da_journal(path)
        journal = sessione.Journal.riprendi(path)
        nome_ripreso = settimana.nome_origine
        print(f"\n  Sessione ripristinata: settimana del {settimana.lunedi.strftime('%d-%m-%Y')}")
        display_menu()
        lunedi = settimana.lunedi
    elif tipo == "xlsx":
        try:
            settimana = sessione.carica_da_xlsx(path)
        except ValueError as e:
            print(f"  ERRORE: {e}")
            sys.exit(1)
        nome_ripreso = settimana.nome_origine
        print(f"\n  Ripreso: {path}")
        display_menu()
        lunedi = chiedi_settimana(settimana.lunedi)
        settimana.lunedi = lunedi
        journal = sessione.Journal(OUTPUT_DIR / sessione.nome_journal(lunedi))
        journal.avvia(settimana)
    else:
        display_menu()
        lunedi = chiedi_settimana()
        settimana = sessione.Settimana(lunedi)
        journal = sessione.Journal(OUTPUT_DIR / sessione.nome_journal(lunedi))
        journal.avvia(settimana)

    stato = {"nome_ripreso": nome_ripreso, "percorso_salvato": None, "uscita_richiesta": False}

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

    salvataggio_ok, cartella_salvata, percorso_finale = menu_uscita_salvataggio(
        settimana, journal, stato["nome_ripreso"], stato["percorso_salvato"]
    )

    if salvataggio_ok and cartella_salvata:
        print("\n  Operazioni aggiuntive:")
        print("    1. Aggiorna riepilogo annuale")
        print("    2. Genera bollettini Word (ITA/ENG)")
        print("    3. Entrambi (annuale + bollettini)")
        print("    Invio = nessuna operazione aggiuntiva")
        risp_extra = input("\n  Scelta: ").strip()
        if risp_extra in ("1", "3"):
            percorso, n = esportatori.esporta_riepilogo_annuale(
                settimana, cartella_salvata, scegli_duplicati=_scegli_duplicati_interattivo)
            if percorso:
                print(f"\n  [OK] Riepilogo annuale aggiornato: {percorso}")
                print(f"       {n} giorni esportati, fattore={settimana.fattore}")
            else:
                print("  Nessun dato da esportare nel riepilogo annuale.")
        if risp_extra in ("2", "3"):
            soglie = esportatori.carica_soglie(cartella_salvata)
            if soglie is None:
                print(f"  ERRORE: file soglie '{esportatori.SOGLIE_FILE_NOME}' non trovato.")
            else:
                creati = esportatori.genera_bollettini_word(settimana, cartella_salvata, soglie=soglie)
                if creati:
                    for p in creati:
                        print(f"  Bollettino: {p}")
                    print("  Bollettini Word generati.")
                else:
                    print("  Nessun dato presente: bollettini Word non generati.")

    print("\n  Sessione terminata.\n")

    if getattr(sys, "frozen", False):
        input("Premi INVIO per chiudere...")


if __name__ == "__main__":
    main()
