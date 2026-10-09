#!/usr/bin/env python3
"""
Dominio puro del Conta Pollinica: codici, date, soglie, concentrazioni,
interpretazione dei comandi e struttura del bollettino.

Nessuna funzione di questo modulo fa I/O (file, stdin, Excel, Word):
per questo e' interamente coperto da test (vedi tests/test_dominio.py).
"""

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timedelta

# ============================================================
# Codici specie
# ============================================================
GIORNI_NOMI = {
    1: "lunedi", 2: "martedi", 3: "mercoledi",
    4: "giovedi", 5: "venerdi", 6: "sabato", 7: "domenica",
}
GIORNI_ABBREV = ["LUN", "MAR", "MER", "GIO", "VEN", "SAB", "DOM"]

CODICI_SPECIE = {
    "01": "ACERACEAE", "02": "ALTRI POLLINI", "03": "BETULACEAE",
    "04": "Alnus", "05": "Betula", "06": "CANNABACEAE",
    "07": "CHENO-AMAR", "08": "COMPOSITAE", "09": "Altre compositae",
    "10": "Ambrosia", "11": "Artemisia", "12": "CORYLACEAE (somma c+o)",
    "13": "Carpinus/Ostrya", "14": "Carpinus", "15": "Ostrya carpinifolia",
    "16": "Corylus avellana", "17": "CUP-TAXACEAE", "18": "ERICACEAE",
    "19": "EUPHORBIACEAE", "20": "FAGACEAE", "21": "Castanea sativa",
    "22": "Fagus sylvatica", "23": "Quercus", "24": "GRAMINEAE",
    "25": "HIPPOCASTANACEAE", "26": "JUGLANDACEAE", "27": "LAURACEAE",
    "28": "MIMOSACEAE", "29": "MORACEAE", "30": "MYRTACEAE",
    "31": "OLEACEAE", "32": "Altre oleaceae", "33": "Fraxinus",
    "34": "Ligustrum", "35": "Olea", "36": "PINACEAE",
    "37": "PLANTAGINACEAE", "38": "PLATANACEAE",
    "39": "POLLINI NON IDENTIFICATI", "40": "POLYGONACEAE",
    "41": "SALICACEAE", "42": "Populus", "43": "Salix",
    "44": "TILIACEAE", "45": "ULMACEAE", "46": "UMBELLIFERAE",
    "47": "URTICACEAE",
    "48": "Alternaria", "49": "Botrytis", "50": "Cladosporium",
    "51": "Curvularia", "52": "Epicoccum", "53": "Helminthosporium",
    "54": "Pithomyces", "55": "Pleospora", "56": "Polythrincium",
    "57": "Stemphylium", "58": "Tetraploa", "59": "Torula",
}

POLLINI_CODICI = [f"{i:02d}" for i in range(1, 48)]    # 47 pollini
SPORE_CODICI = [f"{i:02d}" for i in range(48, 60)]      # 12 spore
TUTTI_CODICI = POLLINI_CODICI + SPORE_CODICI

# Codici evidenziati in verde nei riepiloghi annuali (specie di interesse)
ANNUALE_VERDE_CODICI = {
    "05", "12", "13", "14", "15", "17", "18", "19",
    "22", "23", "24", "33", "35", "36", "47",
}
ANNUALE_VERDE_CHIARO_CODICI = {"48"}
ANNUALE_BOLD_CODICI = ANNUALE_VERDE_CODICI | ANNUALE_VERDE_CHIARO_CODICI

# ============================================================
# Struttura del bollettino pollinico (indipendente dal layout Excel)
# ============================================================
# (nome_ita, nome_eng, codici_da_sommare, famiglia_soglia)
# famiglia_soglia: nome della famiglia nel file soglie (concentrazioni_polliniche.xlsx)
BOLLETTINO_RIGHE = (
    ("ACERACEAE",           "ACERACEAE",             ("01",),                       "Aceracee"),
    ("BETULACEAE",          "BETULACEAE",             ("03", "04", "05"),            "Betulaceae"),
    ("Alnus",               "Alnus",                  ("04",),                       "Betulaceae"),
    ("Betula",               "Betula",                ("05",),                       "Betulaceae"),
    ("CHENO-AMAR",          "CHENO-AMAR",             ("07",),                       "Cheno-Amarantaceae"),
    ("COMPOSITAE",          "COMPOSITAE",             ("08", "09", "10", "11"),      "Composite"),
    ("Altre compositae",    "Other composites",       ("09",),                       "Composite"),
    ("Ambrosia",            "Ragweed",                ("10",),                       "Composite"),
    ("Artemisia",           "Mugwort",                ("11",),                       "Composite"),
    ("CORYLACEAE",          "CORYLACEAE",             ("12", "13", "14", "15", "16"), "Corilacee"),
    ("Carpinus/Ostrya",     "Hornbeam/Hop hornbeam",  ("13",),                       "Corilacee"),
    ("Carpinus",            "Hornbeam",               ("14",),                       "Corilacee"),
    ("Ostrya carpinifolia", "Hop hornbeam",           ("15",),                       "Corilacee"),
    ("Corylus avellana",    "Hazel",                  ("16",),                       "Corilacee"),
    ("CUP-TAXACEAE",        "Cup-Taxaceae",           ("17",),                       "Cupressaceae + Taxaceae"),
    ("FAGACEAE",            "FAGACEAE",               ("20", "21", "22", "23"),      "Fagaceae"),
    ("Castanea sativa",     "Sweet chestnut",         ("21",),                       "Fagaceae"),
    ("Fagus sylvatica",     "Beech",                  ("22",),                       "Fagaceae"),
    ("Quercus",             "Oak",                    ("23",),                       "Fagaceae"),
    ("GRAMINEAE",           "Grasses",                ("24",),                       "Graminaceae"),
    ("OLEACEAE",            "OLEACEAE",               ("31", "32", "33", "34", "35"), "Oleaceae"),
    ("Altre oleaceae",      "Other olive family",     ("32",),                       "Oleaceae"),
    ("Fraxinus",            "Ash",                    ("33",),                       "Oleaceae"),
    ("Ligustrum",           "Privet",                 ("34",),                       "Oleaceae"),
    ("Olea",                "Olive",                  ("35",),                       "Oleaceae"),
    ("PINACEAE",            "PINACEAE",               ("36",),                       "Pinaceae"),
    ("PLANTAGINACEAE",      "PLANTAGINACEAE",         ("37",),                       "Plantaginaceae"),
    ("PLATANACEAE",         "Plane tree",             ("38",),                       "Platanaceae"),
    ("SALICACEAE",          "SALICACEAE",             ("41", "42", "43"),            "Salicaceae"),
    ("Populus",             "Poplar",                 ("42",),                       "Salicaceae"),
    ("Salix",                "Willow",                ("43",),                       "Salicaceae"),
    ("ULMACEAE",            "Elm family",             ("45",),                       "Ulmaceae"),
    ("URTICACEAE",          "Nettle",                 ("47",),                       "Urticaceae"),
    ("Alternaria",          "Alternaria",             ("48",),                       "Alternaria"),
    ("Cladosporium",        "Cladosporium",           ("50",),                       "Cladosporium"),
)

# Soglie di fallback (usate solo se concentrazioni_polliniche.xlsx non e' leggibile)
SOGLIE_FALLBACK = {
    "Aceracee": (0.9, 19.9, 39.9),
    "Betulaceae": (0.5, 15.9, 49.9),
    "Cheno-Amarantaceae": (0.0, 4.9, 24.9),
    "Composite": (0.0, 4.9, 24.9),
    "Corilacee": (0.5, 15.9, 49.9),
    "Cupressaceae + Taxaceae": (3.9, 29.9, 89.9),
    "Fagaceae": (0.9, 19.9, 39.9),
    "Graminaceae": (0.5, 9.9, 29.9),
    "Oleaceae": (0.5, 4.9, 24.9),
    "Pinaceae": (0.9, 14.9, 49.9),
    "Plantaginaceae": (0.0, 0.4, 1.9),
    "Platanaceae": (0.9, 19.9, 39.9),
    "Salicaceae": (0.9, 19.9, 39.9),
    "Ulmaceae": (0.9, 19.9, 39.9),
    "Urticaceae": (1.9, 19.9, 69.9),
    "Alternaria": (1.9, 19.0, 100.0),
    "Cladosporium": (99.9, 499.0, 1000.0),
}

FATTORE_DEFAULT = 0.4

# ============================================================
# Costanti bollettino Word
# ============================================================
MESI_ITA = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
            "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
GIORNI_ITA_LONG = ["Luned\xec", "Marted\xec", "Mercoled\xec", "Gioved\xec",
                    "Venerd\xec", "Sabato", "Domenica"]
GIORNI_ENG_LONG = ["Monday", "Tuesday", "Wednesday", "Thursday",
                    "Friday", "Saturday", "Sunday"]


# ============================================================
# Codici e quantita'
# ============================================================
def normalizza_codice(codice):
    """Normalizza un codice a 2 cifre (es. '5' -> '05')."""
    if codice.isdigit() and len(codice) == 1:
        return codice.zfill(2)
    return codice


def codice_valido(codice):
    return codice in CODICI_SPECIE


# ============================================================
# Layout del foglio Excel 'riepilogo_settimana' (fisso, non modificare:
# vedi CLAUDE.md — pollini in righe 6-52, spore in 58-69)
# ============================================================
def codice_to_row(codice):
    """Codice specie -> riga nel foglio riepilogo_settimana."""
    n = int(codice)
    if 1 <= n <= 47:          # Pollini
        return n + 5
    if 48 <= n <= 59:         # Spore
        return n + 10
    return None


def giorno_to_col(giorno_num):
    """Giorno (1=lun, 7=dom) -> colonna dati grezzi (G=7 ... M=13)."""
    return giorno_num + 6


# ============================================================
# Date
# ============================================================
_MESI_TXT = {
    "gen": 1, "gennaio": 1, "feb": 2, "febbraio": 2,
    "mar": 3, "marzo": 3, "apr": 4, "aprile": 4,
    "mag": 5, "maggio": 5, "giu": 6, "giugno": 6,
    "lug": 7, "luglio": 7, "ago": 8, "agosto": 8,
    "set": 9, "settembre": 9, "ott": 10, "ottobre": 10,
    "nov": 11, "novembre": 11, "dic": 12, "dicembre": 12,
}


def parse_data_flessibile(testo):
    """Cerca di estrarre una data da testo libero.

    Formati riconosciuti (giorno e mese anche a 1 cifra):
      9-2-2026    9/2/2026    09-02-2026    09/02/2026
      9-2-26      9/2/26      09-02-26      09/02/26
      9 febbraio 2026   9 feb 2026   (mesi italiani)
    Ritorna un datetime oppure None.
    """
    testo = testo.strip()

    match = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})", testo)
    if match:
        g, m, a = match.group(1), match.group(2), match.group(3)
        if len(a) == 2:
            a = "20" + a
        try:
            return datetime(int(a), int(m), int(g))
        except ValueError:
            pass

    match = re.search(r"(\d{1,2})\s+([a-zA-Z]+)\s+(\d{2,4})", testo)
    if match:
        g = int(match.group(1))
        m_txt = match.group(2).lower()
        a = match.group(3)
        if len(a) == 2:
            a = "20" + a
        m = _MESI_TXT.get(m_txt)
        if m:
            try:
                return datetime(int(a), m, g)
            except ValueError:
                pass

    return None


def lunedi_di(dt):
    """Ritorna il lunedi' della settimana che contiene dt."""
    return dt - timedelta(days=dt.weekday())


# ============================================================
# Concentrazione e soglie
# ============================================================
def concentrazione(conta, fattore):
    """Unico punto di calcolo conta -> concentrazione (p/m3). Ritorna float
    arrotondato a 1 decimale, o None se la conta e' 0."""
    if not conta:
        return None
    return round(conta * fattore, 1)


def parse_soglia_max(testo):
    """Estrae il valore massimo da un range come '0 - 0,5', '< 1', '> 50'.
    Ritorna un float."""
    if not testo:
        return 0.0
    testo = str(testo).strip().replace(",", ".")
    if testo.startswith(">"):
        return float("inf")
    m = re.match(r"<\s*([\d.]+)", testo)
    if m:
        return float(m.group(1)) - 0.1
    m = re.search(r"([\d.]+)\s*[-–]\s*([\d.]+)", testo)
    if m:
        return float(m.group(2))
    m = re.match(r"([\d.]+)", testo)
    if m:
        return float(m.group(1))
    return 0.0


def parse_soglie(righe):
    """righe: iterabile di (nome, testo_assente, testo_bassa, testo_media).
    Ritorna dict {nome: (max_assente, max_bassa, max_media)}."""
    soglie = {}
    for nome, t_assente, t_bassa, t_media in righe:
        if not nome or not isinstance(nome, str):
            continue
        nome = nome.strip()
        if not t_assente or not any(c.isdigit() for c in str(t_assente)):
            continue
        soglie[nome] = (
            parse_soglia_max(t_assente),
            parse_soglia_max(t_bassa),
            parse_soglia_max(t_media),
        )
    return soglie


def livello(valore, soglia_tuple):
    """Ritorna il livello ('assente','bassa','media','alta') per un valore p/m3."""
    max_ass, max_bas, max_med = soglia_tuple
    if valore is None or valore <= max_ass:
        return "assente"
    if valore <= max_bas:
        return "bassa"
    if valore <= max_med:
        return "media"
    return "alta"


LIVELLO_COLORE_WORD = {
    "assente": "92D050",
    "bassa": "FFFF00",
    "media": "FFC000",
    "alta": "C00000",
}
# Colori (bg, fg) per il livello nella GUI (stessa palette del template Excel)
LIVELLO_COLORE_GUI = {
    "assente": ("#00B050", "#FFFFFF"),
    "bassa":   ("#FFD966", "#000000"),
    "media":   ("#F4B084", "#000000"),
    "alta":    ("#FF0000", "#FFFFFF"),
}


@dataclass
class RigaBollettino:
    nome_ita: str
    nome_eng: str
    concentrazioni: list       # 7 valori (float)
    livelli_giorno: list       # 7 livelli, uno per concentrazione dello stesso giorno
    media: float
    livello_media: str


def righe_bollettino(conteggi, fattore, soglie):
    """Calcola le righe del bollettino a partire dai conteggi grezzi.

    conteggi: {codice: [v_lun..v_dom]} (conta grezza per giorno)
    soglie: {famiglia: (max_assente, max_bassa, max_media)}
    Ritorna la lista di RigaBollettino per le specie con almeno un dato > 0,
    nell'ordine di BOLLETTINO_RIGHE (identico a quello del bollettino Word).

    Ogni giorno e' colorato secondo la propria concentrazione (non secondo
    la media settimanale): e' l'unico punto di calcolo del colore, usato sia
    dall'anteprima GUI sia dal bollettino Word, cosi' i due non possono piu'
    divergere (difetto #3 della revisione).
    """
    righe = []
    for nome_ita, nome_eng, codici, famiglia in BOLLETTINO_RIGHE:
        totali_giorno = [0] * 7
        for codice in codici:
            vals = conteggi.get(codice, [0] * 7)
            for i in range(7):
                totali_giorno[i] += vals[i]
        if not any(v > 0 for v in totali_giorno):
            continue
        conc = [concentrazione(v, fattore) or 0.0 for v in totali_giorno]
        media = round(sum(conc) / 7, 1)
        soglia_tuple = soglie.get(famiglia, SOGLIE_FALLBACK.get(famiglia, (0.9, 19.9, 39.9)))
        righe.append(RigaBollettino(
            nome_ita=nome_ita, nome_eng=nome_eng,
            concentrazioni=conc,
            livelli_giorno=[livello(c, soglia_tuple) for c in conc],
            media=media,
            livello_media=livello(media, soglia_tuple),
        ))
    return righe


# ============================================================
# Interpretazione dei comandi da tastiera
# ============================================================
@dataclass
class Inserisci:
    codice: str
    quantita: int


class Ripeti:
    """Comando '.': ripete l'ultimo codice inserito."""


@dataclass
class Azione:
    """Comando a singola lettera: h, r, w, l, b, c, n, s, u, d, q."""
    lettera: str


@dataclass
class ComandoNonValido:
    messaggio: str


_RE_QUANTITA = re.compile(r"^(\d{1,2})[xX*](\d+)$")
_LETTERE_AZIONE = set("hrwlbcnsudq")


def interpreta_comando(testo):
    """Interpreta l'input utente durante una sessione giorno.
    Ritorna Inserisci, Ripeti, Azione o ComandoNonValido."""
    testo = testo.strip()
    if not testo:
        return ComandoNonValido("")

    cmd = testo.lower()
    if cmd == ".":
        return Ripeti()
    if cmd in _LETTERE_AZIONE and len(cmd) == 1:
        return Azione(cmd)

    quantita = 1
    codice = testo
    match = _RE_QUANTITA.match(testo)
    if match:
        codice = match.group(1)
        quantita = int(match.group(2))
        if quantita < 1 or quantita > 100:
            return ComandoNonValido(f"Quantita' non valida (1-100).")

    codice = normalizza_codice(codice)
    if not codice_valido(codice):
        return ComandoNonValido(f"Codice non riconosciuto: {testo}")

    return Inserisci(codice=codice, quantita=quantita)


# ============================================================
# Interpretazione dei comandi vocali
#
# La voce e' una seconda sorgente di comandi: produce gli stessi oggetti
# di interpreta_comando() (Inserisci, Ripeti, Azione, ComandoNonValido).
# Solo azioni reversibili sono raggiungibili a voce: salvare, chiudere la
# giornata o uscire restano da tastiera/pulsante, cosi' un falso
# riconoscimento non puo' chiudere o perdere nulla.
#
# Parola di attivazione (come "Count" in EcoCount, Allen & Sewell 2014): ogni
# frase deve iniziare con "conta" ("conta ventiquattro", "conta annulla").
# La grammatica chiusa da sola non basta: "due", "tre", "sei" sono parole
# comuni e una conversazione in laboratorio le registrerebbe come granuli.
# Disattivabile con la chiave "parola_attivazione" di pollencounter.cfg.
# ============================================================
PAROLA_ATTIVAZIONE_DEFAULT = "conta"
QUANTITA_VOCALE_MIN = 2
QUANTITA_VOCALE_MAX = 20

_UNITA = ["", "uno", "due", "tre", "quattro", "cinque", "sei", "sette",
          "otto", "nove", "dieci", "undici", "dodici", "tredici",
          "quattordici", "quindici", "sedici", "diciassette", "diciotto",
          "diciannove"]
_DECINE = {2: "venti", 3: "trenta", 4: "quaranta", 5: "cinquanta"}

# Forme parlate aggiuntive rispetto a codice (a numero) e nome in CODICI_SPECIE.
# Gia' normalizzate (minuscolo, senza accenti). Estendibili dall'utente con la
# chiave "sinonimi_vocali" di pollencounter.cfg (vedi sessione.leggi_sinonimi_vocali).
SINONIMI_VOCALI_DEFAULT = {
    "01": ["aceracee", "acero", "aceri"],
    "03": ["betulacee"],
    "04": ["ontano"],
    "05": ["betulla"],
    "06": ["cannabacee", "canapa", "luppolo"],
    "07": ["cheno amarantacee", "amaranto"],
    "08": ["composite", "asteracee"],
    "09": ["altre composite"],
    "12": ["corilacee"],
    "14": ["carpino bianco"],
    "15": ["carpino nero"],
    "16": ["nocciolo"],
    "17": ["cupressacee taxacee", "cipresso", "cipressi"],
    "18": ["erica"],
    "20": ["fagacee"],
    "21": ["castagno"],
    "22": ["faggio"],
    "23": ["quercia"],
    "24": ["graminacee"],
    "25": ["ippocastano"],
    "26": ["noce"],
    "27": ["alloro"],
    "28": ["mimosa"],
    "29": ["moracee", "gelso"],
    "30": ["mirto", "eucalipto"],
    "31": ["oleacee"],
    "32": ["altre oleacee"],
    "33": ["frassino"],
    "34": ["ligustro"],
    "35": ["olivo", "ulivo"],
    "36": ["pinacee", "pino", "pini"],
    "37": ["plantaginacee"],
    "38": ["platanacee", "platano", "platani"],
    "39": ["non identificati"],
    "40": ["poligonacee", "acetosa"],
    "41": ["salicacee"],
    "42": ["pioppo"],
    "43": ["salice"],
    "44": ["tiliacee", "tiglio"],
    "45": ["ulmacee", "olmo"],
    "46": ["ombrellifere"],
    "47": ["urticacee", "ortica"],
}

class Totale:
    """Comando vocale 'totale': il PC legge il totale del giorno (sola lettura)."""


COMANDI_VOCALI = {
    "annulla": Azione("u"),
    "ripeti": Ripeti(),
    "ancora": Ripeti(),
    "totale": Totale(),
}


def numero_in_parole(n):
    """Numero 1-59 in lettere italiane, senza accenti ('tre' -> 'ventitre')."""
    if not 1 <= n <= 59:
        raise ValueError(f"Numero fuori intervallo: {n}")
    if n < 20:
        return _UNITA[n]
    base = _DECINE[n // 10]
    unita = n % 10
    if unita == 0:
        return base
    if unita in (1, 8):               # elisione: ventuno, ventotto
        base = base[:-1]
    return base + _UNITA[unita]


def normalizza_parlato(testo):
    """Minuscolo, senza accenti ne' punteggiatura, spazi singoli."""
    testo = unicodedata.normalize("NFD", testo.lower())
    testo = "".join(c for c in testo if not unicodedata.combining(c))
    testo = re.sub(r"\([^)]*\)", " ", testo)      # '(somma c+o)' non si pronuncia
    testo = re.sub(r"[^a-z0-9]+", " ", testo)
    return testo.strip()


def costruisci_vocabolario_vocale(extra=None, attivazione=PAROLA_ATTIVAZIONE_DEFAULT):
    """Ritorna (vocabolario, avvisi): vocabolario = {frase parlata: codice}.

    Per ogni codice: il numero in lettere, il nome in CODICI_SPECIE e i
    sinonimi (default + 'extra' dell'utente). Una frase che porterebbe a due
    codici diversi viene scartata dalla seconda comparsa e segnalata in
    'avvisi' (testi in italiano, mostrabili all'utente)."""
    vocabolario = {}
    avvisi = []

    def aggiungi(frase, codice, origine):
        frase = normalizza_parlato(frase)
        if not frase:
            return
        if frase in COMANDI_VOCALI or frase == "per" or (attivazione and frase == attivazione):
            avvisi.append(f"'{frase}' ({origine}) e' una parola riservata della voce: ignorata.")
            return
        esistente = vocabolario.get(frase)
        if esistente is not None and esistente != codice:
            avvisi.append(f"'{frase}' indica sia {esistente} sia {codice}: "
                          f"tengo {esistente}, ignoro {codice}.")
            return
        vocabolario[frase] = codice

    for codice, nome in CODICI_SPECIE.items():
        aggiungi(numero_in_parole(int(codice)), codice, "numero")
        aggiungi(nome, codice, "nome")
    for sorgente, origine in ((SINONIMI_VOCALI_DEFAULT, "sinonimo"), (extra or {}, "sinonimo personale")):
        for codice, forme in sorgente.items():
            codice = normalizza_codice(str(codice))
            if not codice_valido(codice):
                avvisi.append(f"Sinonimi per un codice inesistente: {codice}.")
                continue
            for forma in forme:
                aggiungi(forma, codice, origine)
    return vocabolario, avvisi


def costruisci_grammatica(vocabolario, attivazione=PAROLA_ATTIVAZIONE_DEFAULT):
    """Elenco di tutte le frasi che il riconoscitore puo' restituire, ciascuna
    preceduta dalla parola di attivazione (se non vuota)."""
    frasi = set(COMANDI_VOCALI)
    for frase in vocabolario:
        frasi.add(frase)
        for q in range(QUANTITA_VOCALE_MIN, QUANTITA_VOCALE_MAX + 1):
            frasi.add(f"{frase} per {numero_in_parole(q)}")
    if attivazione:
        frasi = {f"{attivazione} {f}" for f in frasi}
    return sorted(frasi)


def _quantita_da_parole(testo):
    for q in range(QUANTITA_VOCALE_MIN, QUANTITA_VOCALE_MAX + 1):
        if numero_in_parole(q) == testo:
            return q
    return None


def interpreta_vocale(testo, vocabolario, attivazione=PAROLA_ATTIVAZIONE_DEFAULT):
    """Interpreta una frase riconosciuta dalla voce. Stessi tipi di
    interpreta_comando() piu' Totale. Il testo fuori vocabolario o senza la
    parola di attivazione ('[unk]', rumore di fondo) da' ComandoNonValido con
    messaggio vuoto: va scartato in silenzio."""
    if "[unk]" in testo.lower():
        return ComandoNonValido("")
    frase = normalizza_parlato(testo)
    if attivazione:
        prefisso = attivazione + " "
        if not frase.startswith(prefisso):
            return ComandoNonValido("")
        frase = frase[len(prefisso):].strip()
    if not frase:
        return ComandoNonValido("")
    if frase in COMANDI_VOCALI:
        return COMANDI_VOCALI[frase]

    quantita = 1
    if " per " in frase:
        frase, _, parole_q = frase.rpartition(" per ")
        quantita = _quantita_da_parole(parole_q)
        if quantita is None:
            return ComandoNonValido(f"Quantita' non capita: {testo}")

    codice = vocabolario.get(frase)
    if codice is None:
        return ComandoNonValido("")
    return Inserisci(codice=codice, quantita=quantita)
