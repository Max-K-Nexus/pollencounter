#!/usr/bin/env python3
"""
Archivio: struttura della cartella di lavoro e nomi dei file.

Unico punto del programma che sa *dove* vive un file e *come* si chiama.
Chi scrive i file (esportatori.py) riceve il percorso gia' fatto; chi lo
calcola (GUI e CLI) lo chiede qui. Nessun altro modulo deve costruire nomi
di file con una f-string.

Struttura:

    <radice>/
    |- Cheatsheet_Codici_Specie.docx      elenco codici stampabile
    |- concentrazioni_polliniche.xlsx     soglie personalizzate (opzionale)
    |- 2026/
    |  |- Riepilogo_Annuale_2026.xlsx
    |  |- S41_2026-10-05/
    |  |  |- Conta_Pollinica_05-10-2026.xlsx
    |  |  |- Bollettino_ITA_05-10-2026.docx
    |  |  +- Bollettino_ENG_05-10-2026.docx
    |  +- S42_2026-10-12/
    +- in_corso/
       +- ~sessione_19-10-2026.jsonl      lavori non salvati, di ogni anno

I lavori non salvati stanno tutti in 'in_corso/' e non sotto l'anno: cosi'
restano visibili anche a cavallo di due stagioni.

Nessun I/O oltre a mkdir/stat/glob/move: niente openpyxl (la posizione di un
file dice gia' che cos'e', non serve aprirlo), niente dialoghi, niente print.
"""

import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import dominio


# ============================================================
# Nomi (unica fonte)
# ============================================================
PREFISSO_SETTIMANALE = "Conta_Pollinica_"
PREFISSO_ANNUALE = "Riepilogo_Annuale_"
PREFISSO_BOLLETTINO = "Bollettino_"
PREFISSO_JOURNAL = "~sessione_"
NOME_ELENCO_CODICI = "Cheatsheet_Codici_Specie.docx"

# File di supporto che possono trovarsi nella radice e non appartengono a una
# settimana: il riordino non li tocca mai. Le soglie, in particolare, vengono
# lette proprio dalla radice (esportatori.carica_soglie).
NOME_SOGLIE = "concentrazioni_polliniche.xlsx"
NOME_TEMPLATE = "Polline_Template_Settimanale.xlsx"

NOME_IN_CORSO = "in_corso"
NOME_DA_CONTROLLARE = "da_controllare"
PREFISSO_RESOCONTO = "_riordino_"

# Formato data usato in tutti i nomi di file (quello storico del programma).
FORMATO_DATA = "%d-%m-%Y"


def nome_settimanale(lunedi):
    return f"{PREFISSO_SETTIMANALE}{lunedi.strftime(FORMATO_DATA)}.xlsx"


def nome_annuale(anno):
    return f"{PREFISSO_ANNUALE}{anno}.xlsx"


def nome_bollettino(lingua, lunedi):
    """lingua: 'ITA' o 'ENG'."""
    return f"{PREFISSO_BOLLETTINO}{lingua}_{lunedi.strftime(FORMATO_DATA)}.docx"


def nome_journal(lunedi):
    return f"{PREFISSO_JOURNAL}{lunedi.strftime(FORMATO_DATA)}.jsonl"


def nome_cartella_settimana(lunedi):
    """'S42_2026-10-12': settimana ISO (la stessa dei fogli 'Wnn' del riepilogo
    annuale) e data del lunedi'. In ordine alfabetico = in ordine cronologico."""
    return f"S{dominio.numero_settimana(lunedi):02d}_{lunedi.strftime('%Y-%m-%d')}"


# ============================================================
# Percorsi
# ============================================================
def cartella_anno(radice, anno):
    return Path(radice) / str(anno)


def cartella_settimana(radice, lunedi):
    """L'anno e' quello civile del lunedi', lo stesso del riepilogo annuale.
    (Il numero ISO puo' appartenere all'anno successivo: lunedi' 29-12-2025 e'
    la settimana 1, e sta in 2025/S01_2025-12-29.)"""
    return cartella_anno(radice, lunedi.year) / nome_cartella_settimana(lunedi)


def percorso_settimanale(radice, lunedi):
    return cartella_settimana(radice, lunedi) / nome_settimanale(lunedi)


def percorso_annuale(radice, anno):
    return cartella_anno(radice, anno) / nome_annuale(anno)


def percorso_bollettino(radice, lingua, lunedi):
    return cartella_settimana(radice, lunedi) / nome_bollettino(lingua, lunedi)


def cartella_in_corso(radice):
    return Path(radice) / NOME_IN_CORSO


def percorso_journal(radice, lunedi):
    return cartella_in_corso(radice) / nome_journal(lunedi)


def percorso_elenco_codici(radice):
    return Path(radice) / NOME_ELENCO_CODICI


# ── creazione delle cartelle (gli unici mkdir dell'archivio) ──
def prepara_anno(radice, anno):
    cartella = cartella_anno(radice, anno)
    cartella.mkdir(parents=True, exist_ok=True)
    return cartella


def prepara_settimana(radice, lunedi):
    cartella = cartella_settimana(radice, lunedi)
    cartella.mkdir(parents=True, exist_ok=True)
    return cartella


def prepara_in_corso(radice):
    cartella = cartella_in_corso(radice)
    cartella.mkdir(parents=True, exist_ok=True)
    return cartella


# ============================================================
# Lettura dei nomi
# ============================================================
_RE_DATA_NOME = re.compile(r"(\d{2})-(\d{2})-(\d{4})")
_RE_CARTELLA_SETTIMANA = re.compile(r"^S(\d{2})_(\d{4})-(\d{2})-(\d{2})$")
_RE_ANNO = re.compile(r"^\d{4}$")


def legge_lunedi_da_nome(nome):
    """Ricava il lunedi' della settimana dal nome di un file o di una cartella
    ('S42_2026-10-12', 'Conta_Pollinica_12-10-2026.xlsx',
    '~sessione_12-10-2026.jsonl'), o None se il nome non lo contiene.

    La data trovata viene sempre normalizzata al lunedi' della sua settimana:
    un file nominato con un giorno qualsiasi resta attribuito alla settimana
    giusta."""
    nome = Path(nome).name
    m = _RE_CARTELLA_SETTIMANA.match(nome)
    if m:
        _, anno, mese, giorno = m.groups()
        return _data_valida(int(anno), int(mese), int(giorno))
    m = _RE_DATA_NOME.search(nome)
    if m:
        giorno, mese, anno = m.groups()
        return _data_valida(int(anno), int(mese), int(giorno))
    return None


def _data_valida(anno, mese, giorno):
    try:
        return dominio.lunedi_di(datetime(anno, mese, giorno))
    except ValueError:
        return None


def e_settimanale(nome):
    return Path(nome).name.startswith(PREFISSO_SETTIMANALE) and nome.endswith(".xlsx")


def e_annuale(nome):
    return Path(nome).name.startswith(PREFISSO_ANNUALE) and nome.endswith(".xlsx")


def e_bollettino(nome):
    return Path(nome).name.startswith(PREFISSO_BOLLETTINO) and nome.endswith(".docx")


def e_journal(nome):
    return Path(nome).name.startswith(PREFISSO_JOURNAL) and nome.endswith(".jsonl")


# ============================================================
# Elenco dell'archivio
# ============================================================
@dataclass
class VoceArchivio:
    """Una settimana dell'archivio: tutto cio' che esiste su disco per quel
    periodo, in un'unica voce (prima journal e .xlsx erano due righe scollegate
    nell'elenco iniziale)."""
    lunedi: datetime = None         # None = file che non dichiara la settimana nel nome
    cartella: Path = None
    xlsx: Path = None
    journal: Path = None
    bollettini: list = field(default_factory=list)
    mtime: float = 0.0
    fuori_struttura: bool = False   # non e' ancora nell'albero anno/settimana

    @property
    def stato(self):
        return "non_salvato" if self.journal else "salvato"

    @property
    def etichetta(self):
        """Come chiamare la voce nell'elenco, gia' pronta da mostrare: il
        periodo se lo conosciamo, altrimenti il nome del file (fuori schema, da
        riordinare a mano) che va riportato tale e quale."""
        if self.lunedi:
            descrizione = dominio.descrivi_settimana(self.lunedi)
            return descrizione[0].upper() + descrizione[1:]
        return self.xlsx.name if self.xlsx else ""


def elenca(radice, legacy_xlsx=()):
    """Tutte le settimane presenti in 'radice': lavori non salvati per primi,
    poi le settimane piu' recenti.

    Legge soltanto nomi e date di modifica: non apre nessun file.
    'legacy_xlsx' sono file .xlsx della vecchia cartella piatta il cui nome non
    dice la settimana (es. 'rifatto_ContaPoll_160226.xlsx'), gia' riconosciuti
    come settimanali da chi chiama (sessione.cerca_file_ripresa): qui entrano
    come voci da riordinare."""
    radice = Path(radice)
    voci = {}

    def voce(chiave, lunedi=None, fuori_struttura=False):
        if chiave not in voci:
            voci[chiave] = VoceArchivio(lunedi=lunedi, fuori_struttura=fuori_struttura)
        v = voci[chiave]
        if fuori_struttura:
            v.fuori_struttura = True
        return v

    # ── cartelle anno/settimana (la struttura nuova) ──
    for cartella in _sottocartelle_settimana(radice):
        lunedi = legge_lunedi_da_nome(cartella.name)
        if lunedi is None:
            continue
        v = voce(lunedi.date(), lunedi)
        v.cartella = cartella
        for f in sorted(cartella.iterdir()):
            if not f.is_file():
                continue
            if e_settimanale(f.name) and v.xlsx is None:
                v.xlsx = f
            elif e_bollettino(f.name):
                v.bollettini.append(f)
            _aggiorna_mtime(v, f)

    # ── file settimanali rimasti nella radice piatta ──
    for f in _file_diretti(radice, ".xlsx"):
        if not e_settimanale(f.name):
            continue
        lunedi = legge_lunedi_da_nome(f.name)
        if lunedi is None:
            continue
        v = voce(lunedi.date(), lunedi, fuori_struttura=True)
        if v.xlsx is None:
            v.xlsx = f
        _aggiorna_mtime(v, f)

    # ── .xlsx messi da parte dal riordino: restano raggiungibili dall'elenco ──
    for f in _file_diretti(Path(radice) / NOME_DA_CONTROLLARE, ".xlsx"):
        lunedi = legge_lunedi_da_nome(f.name)
        chiave = lunedi.date() if lunedi else ("file", f.name)
        v = voce(chiave, lunedi, fuori_struttura=True)
        if v.xlsx is None:
            v.xlsx = f
        _aggiorna_mtime(v, f)

    # ── .xlsx fuori schema, riconosciuti da chi chiama ──
    for f in legacy_xlsx:
        f = Path(f)
        if not f.is_file() or f.parent != radice or e_settimanale(f.name) or e_annuale(f.name):
            continue
        v = voce(("file", f.name), None, fuori_struttura=True)
        v.xlsx = f
        _aggiorna_mtime(v, f)

    # ── lavori non salvati: in_corso/ e (vecchia posizione) radice ──
    for f in sorted(cartella_in_corso(radice).glob(f"{PREFISSO_JOURNAL}*.jsonl")) \
            + sorted(radice.glob(f"{PREFISSO_JOURNAL}*.jsonl")):
        lunedi = legge_lunedi_da_nome(f.name)
        if lunedi is None:
            continue
        v = voce(lunedi.date(), lunedi)
        if v.journal is None:                   # in_corso/ ha la precedenza
            v.journal = f
        _aggiorna_mtime(v, f)

    # Una cartella di settimana vuota (creata e mai riempita, p.es. dopo un
    # salvataggio non riuscito) non e' una voce dell'elenco.
    reali = [v for v in voci.values() if v.xlsx or v.journal]
    return sorted(reali, key=_ordine_elenco)


def _ordine_elenco(v):
    """Prima i lavori non salvati, poi le settimane piu' recenti; le voci senza
    settimana (file fuori schema) in fondo, per data di modifica."""
    giorni = v.lunedi.toordinal() if v.lunedi else 0
    return (0 if v.journal else 1, -giorni, -v.mtime)


def _aggiorna_mtime(voce, path):
    try:
        voce.mtime = max(voce.mtime, path.stat().st_mtime)
    except OSError:
        pass


def _sottocartelle_settimana(radice):
    """Le cartelle <radice>/<anno>/S<nn>_<data>/ esistenti."""
    try:
        anni = sorted(d for d in radice.iterdir() if d.is_dir() and _RE_ANNO.match(d.name))
    except OSError:
        return []
    trovate = []
    for anno in anni:
        try:
            trovate.extend(sorted(d for d in anno.iterdir()
                                  if d.is_dir() and _RE_CARTELLA_SETTIMANA.match(d.name)))
        except OSError:
            continue
    return trovate


def _file_diretti(radice, *suffissi):
    try:
        return sorted(f for f in radice.iterdir()
                      if f.is_file() and f.suffix.lower() in suffissi)
    except OSError:
        return []


# ============================================================
# Riordino dei file della vecchia cartella piatta
# ============================================================
def piano_riordino(radice):
    """Elenco [(origine, destinazione)] per portare nella struttura i file
    rimasti nella radice. Non tocca nulla: lo spostamento e' esegui_riordino().

    Guarda solo .xlsx/.docx/.jsonl rimasti al primo livello: tutto il resto
    (soglie personalizzate, cfg, file dell'utente) resta dov'e'. I file il cui
    nome non dice la settimana finiscono in 'da_controllare/' invece di essere
    attribuiti a indovinare."""
    radice = Path(radice)
    piano = []
    intoccabili = {NOME_ELENCO_CODICI, NOME_SOGLIE, NOME_TEMPLATE}
    for f in _file_diretti(radice, ".xlsx", ".docx", ".jsonl"):
        if f.name in intoccabili or f.name.startswith(PREFISSO_RESOCONTO):
            continue
        destinazione = _destinazione(radice, f)
        if destinazione is not None and destinazione != f:
            piano.append((f, destinazione))
    return piano


def _destinazione(radice, f):
    if e_journal(f.name):
        lunedi = legge_lunedi_da_nome(f.name)
        if lunedi is None:
            return None
        return cartella_in_corso(radice) / f.name
    if e_annuale(f.name):
        m = re.search(r"(\d{4})", f.stem)
        if not m:
            return radice / NOME_DA_CONTROLLARE / f.name
        return cartella_anno(radice, m.group(1)) / f.name
    if e_settimanale(f.name) or e_bollettino(f.name):
        lunedi = legge_lunedi_da_nome(f.name)
        if lunedi is None:
            return radice / NOME_DA_CONTROLLARE / f.name
        return cartella_settimana(radice, lunedi) / f.name
    return radice / NOME_DA_CONTROLLARE / f.name


def archivia_originale(radice, lunedi, origine):
    """Sposta nella cartella della settimana il file .xlsx da cui la sessione e'
    stata ripresa, quando era rimasto nella vecchia cartella piatta.

    Serve a non lasciare in giro una seconda copia con dati piu' vecchi di
    quella appena salvata. Se il nome e' lo stesso del file appena salvato, lo
    prefissa con 'originale_' invece di sovrascriverlo: non si cancella e non si
    sovrascrive mai niente. Ritorna la destinazione, o None se non c'era nulla
    da spostare."""
    if origine is None:
        return None
    origine = Path(origine)
    if not origine.is_file() or origine.parent != Path(radice):
        return None
    cartella = cartella_settimana(radice, lunedi)
    destinazione = cartella / origine.name
    if destinazione.exists():
        destinazione = cartella / f"originale_{origine.name}"
    if destinazione.exists():
        return None
    try:
        cartella.mkdir(parents=True, exist_ok=True)
        shutil.move(str(origine), str(destinazione))
    except OSError:
        return None
    return destinazione


def serve_riordino(radice):
    return bool(piano_riordino(radice))


def esegui_riordino(piano, radice=None):
    """Sposta i file del piano. Non sovrascrive mai: se la destinazione esiste
    gia', la voce finisce fra i falliti e l'originale resta dov'e'.
    Ritorna (spostati, falliti) con falliti = [(origine, motivo)].
    Se 'radice' e' indicata, scrive accanto ai file un resoconto di cosa e'
    stato spostato."""
    spostati, falliti = [], []
    for origine, destinazione in piano:
        origine, destinazione = Path(origine), Path(destinazione)
        if destinazione.exists():
            falliti.append((origine, f"'{destinazione.name}' esiste gia' nella cartella di arrivo"))
            continue
        try:
            destinazione.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(origine), str(destinazione))
        except OSError as e:
            falliti.append((origine, str(e)))
        else:
            spostati.append((origine, destinazione))
    if radice is not None and (spostati or falliti):
        _scrivi_resoconto(Path(radice), spostati, falliti)
    return spostati, falliti


def _scrivi_resoconto(radice, spostati, falliti):
    percorso = radice / f"{PREFISSO_RESOCONTO}{datetime.now().strftime('%Y-%m-%d_%H%M')}.txt"
    righe = [f"Riordino dei file del {datetime.now().strftime('%d/%m/%Y %H:%M')}", ""]
    righe.append(f"Spostati: {len(spostati)}")
    for origine, destinazione in spostati:
        righe.append(f"  {origine.name}  ->  {_relativo(destinazione, radice)}")
    if falliti:
        righe += ["", f"Non spostati: {len(falliti)}"]
        for origine, motivo in falliti:
            righe.append(f"  {origine.name}: {motivo}")
    try:
        percorso.write_text("\n".join(righe) + "\n", encoding="utf-8")
    except OSError:
        pass
    return percorso


def _relativo(percorso, radice):
    try:
        return str(Path(percorso).relative_to(radice))
    except ValueError:
        return str(percorso)
