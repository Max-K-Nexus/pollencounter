#!/usr/bin/env python3
"""
Lettura vocale dei pollini (spunti da EcoCount, Allen & Sewell 2014).

Questo modulo e' l'unico che tocca microfono e altoparlanti. Il significato
delle frasi (codici, nomi, quantita', comandi) sta in dominio.py ed e'
interamente testato li'; qui c'e' solo:

- ``Ascoltatore``: legge blocchi audio, li passa a un riconoscitore e mette
  in coda le frasi capite. Non conosce Vosk: riceve un riconoscitore con
  l'interfaccia di ``vosk.KaldiRecognizer`` (AcceptWaveform/Result/Reset),
  cosi' si puo' provare con un riconoscitore finto e, in futuro, affiancare
  un altro motore (es. SAPI di Windows) senza toccare la GUI.
- ``Sintesi``: il PC ripete a voce cio' che ha capito (talkback di EcoCount),
  cosi' l'operatore non stacca gli occhi dal microscopio.
- ``crea_voce``: assembla motore, microfono e sintesi, oppure spiega in
  italiano cosa manca.

Dipendenze opzionali (come ``sv_ttk``): senza ``vosk``/``sounddevice`` la GUI
funziona normalmente e il pulsante Voce spiega cosa installare.
"""

import json
import queue
import threading
import time
from pathlib import Path

import percorsi

try:
    import vosk
except ImportError:
    vosk = None

try:
    import sounddevice
except (ImportError, OSError):      # OSError: manca la libreria di sistema PortAudio
    sounddevice = None

try:
    import pyttsx3
except ImportError:
    pyttsx3 = None

SAMPLE_RATE = 16000
SOGLIA_CONFIDENZA_DEFAULT = 0.5
CODA_MUTO_SECONDI = 0.4          # coda di silenzio dopo la sintesi (eco nella stanza)
NOME_MODELLO_PREFISSO = "vosk-model"


class VoceNonDisponibile(Exception):
    """Il messaggio e' scritto per l'utente (italiano, senza gergo)."""


# ============================================================
# Estrazione della frase da un risultato del riconoscitore (pura)
# ============================================================
def frase_da_risultato(risultato_json, soglia=SOGLIA_CONFIDENZA_DEFAULT):
    """Da un risultato Vosk (stringa JSON) ritorna la frase capita, oppure ''
    se vuota, contenente '[unk]' o con almeno una parola sotto soglia di
    confidenza (nel dubbio non si registra nulla: meglio ripetere che
    sbagliare specie)."""
    try:
        risultato = json.loads(risultato_json)
    except (TypeError, ValueError):
        return ""
    testo = (risultato.get("text") or "").strip()
    if not testo or "[unk]" in testo:
        return ""
    for parola in risultato.get("result", []):
        if parola.get("conf", 1.0) < soglia:
            return ""
    return testo


# ============================================================
# Ascolto
# ============================================================
class Ascoltatore:
    """Thread di ascolto: blocchi audio -> riconoscitore -> coda di frasi.

    ``sorgente_audio``: iterabile (bloccante) di blocchi ``bytes``.
    ``riconoscitore``: oggetto con AcceptWaveform(bytes)->bool, Result()->str
    (JSON) e Reset().

    Mentre e' "muto" (sintesi vocale in corso) l'audio viene scartato e il
    riconoscitore azzerato: altrimenti il PC sentirebbe se stesso dire
    "graminacee" e registrerebbe un secondo granulo.
    """

    def __init__(self, riconoscitore, sorgente_audio, soglia=SOGLIA_CONFIDENZA_DEFAULT):
        self._riconoscitore = riconoscitore
        self._sorgente = sorgente_audio
        self._soglia = soglia
        self._frasi = queue.Queue()
        self._muto_fino = 0.0
        self._ferma = threading.Event()
        self._thread = None
        self.errore = None

    def avvia(self):
        self._ferma.clear()
        self._thread = threading.Thread(target=self._ciclo, daemon=True)
        self._thread.start()

    def ferma(self):
        self._ferma.set()
        chiudi = getattr(self._sorgente, "close", None)
        if chiudi:
            chiudi()
        if self._thread is not None:
            self._thread.join(timeout=2)

    def silenzia(self, secondi):
        """Ignora l'audio per ``secondi`` (e per la coda di eco)."""
        self._muto_fino = max(self._muto_fino, time.monotonic() + secondi + CODA_MUTO_SECONDI)

    def prossima_frase(self):
        """Prossima frase capita, o None se non ce ne sono (non bloccante)."""
        try:
            return self._frasi.get_nowait()
        except queue.Empty:
            return None

    def _ciclo(self):
        try:
            for blocco in self._sorgente:
                if self._ferma.is_set():
                    break
                if time.monotonic() < self._muto_fino:
                    self._riconoscitore.Reset()
                    continue
                if self._riconoscitore.AcceptWaveform(blocco):
                    frase = frase_da_risultato(self._riconoscitore.Result(), self._soglia)
                    if frase:
                        self._frasi.put(frase)
        except Exception as e:          # microfono scollegato, driver audio, ecc.
            self.errore = str(e)


class SorgenteMicrofono:
    """Blocchi audio dal microfono predefinito (sounddevice)."""

    def __init__(self):
        self._blocchi = queue.Queue()
        self._chiuso = False
        self._stream = sounddevice.RawInputStream(
            samplerate=SAMPLE_RATE, blocksize=4000, dtype="int16", channels=1,
            callback=lambda dati, _n, _t, _s: self._blocchi.put(bytes(dati)))
        self._stream.start()

    def __iter__(self):
        while not self._chiuso:
            try:
                yield self._blocchi.get(timeout=0.2)
            except queue.Empty:
                continue

    def close(self):
        self._chiuso = True
        self._stream.stop()
        self._stream.close()


# ============================================================
# Sintesi vocale (talkback)
# ============================================================
def scegli_voce_italiana(voci):
    """Id della prima voce italiana fra quelle di pyttsx3, o None.
    Riconosce la lingua dichiarata ('it', 'it-IT', anche in bytes con il
    byte di lunghezza iniziale che mette espeak), il nome ('Italian') e gli
    id di espeak ('roa/it'). Senza, la sintesi userebbe la voce predefinita
    (di solito inglese) e storpierebbe i nomi."""
    for voce_ in voci:
        lingue = []
        for lingua in (getattr(voce_, "languages", None) or []):
            if isinstance(lingua, bytes):
                lingua = lingua.decode("utf-8", "ignore")
            lingue.append(str(lingua).strip("\x00\x01\x02\x03\x04\x05").lower())
        nome = (getattr(voce_, "name", "") or "").lower()
        id_ = (getattr(voce_, "id", "") or "").lower()
        if (any(l == "it" or l.startswith(("it-", "it_")) for l in lingue)
                or "italian" in nome or "italiano" in nome
                or id_.endswith("/it") or "italian" in id_):
            return voce_.id
    return None


class Sintesi:
    """Il PC dice a voce cio' che ha capito. Un solo thread possiede il
    motore di sintesi (richiesto da pyttsx3/COM su Windows); le frasi
    arrivano da una coda. Prima di parlare silenzia l'ascoltatore.

    ``avviso``: se non vuoto, la sintesi non e' pienamente disponibile
    (nessuna voce italiana, motore non avviabile): da mostrare all'utente."""

    def __init__(self, ascoltatore=None, durata_stimata=lambda testo: 0.5 + 0.08 * len(testo)):
        self._ascoltatore = ascoltatore
        self._durata = durata_stimata
        self._coda = queue.Queue()
        self.avviso = ""
        self._pronta = threading.Event()
        self._thread = threading.Thread(target=self._ciclo, daemon=True)
        self._thread.start()
        self._pronta.wait(timeout=10)

    def parla(self, testo):
        self._coda.put(testo)

    def ferma(self):
        self._coda.put(None)

    def _ciclo(self):
        motore = None
        try:
            motore = pyttsx3.init()
            id_voce = scegli_voce_italiana(motore.getProperty("voices"))
            if id_voce:
                motore.setProperty("voice", id_voce)
            else:
                self.avviso = ("Nessuna voce italiana installata: il computer pronuncera' "
                               "i nomi con una voce straniera.")
        except Exception as e:
            self.avviso = f"Sintesi vocale non disponibile ({e}): niente conferma parlata."
        finally:
            self._pronta.set()
        while True:
            testo = self._coda.get()
            if testo is None:
                break
            if motore is None:
                continue
            try:
                if self._ascoltatore is not None:
                    self._ascoltatore.silenzia(self._durata(testo))
                motore.say(testo)
                motore.runAndWait()
            except Exception as e:      # un errore di sintesi non deve fermare la conta
                self.avviso = f"Errore della sintesi vocale: {e}"


# ============================================================
# Modello e vocabolario del riconoscitore
# ============================================================
def trova_modello():
    """Cartella del modello Vosk italiano, o None. Cerca in ``modelli/``
    accanto ai sorgenti/bundle e accanto all'eseguibile."""
    for base in (percorsi.BUNDLE_DIR, percorsi.SCRIPT_DIR, percorsi.EXE_DIR):
        cartella = Path(base) / "modelli"
        if cartella.is_dir():
            for candidato in sorted(cartella.iterdir()):
                if candidato.is_dir() and candidato.name.startswith(NOME_MODELLO_PREFISSO):
                    return candidato
    return None


def parole_fuori_modello(grammatica, conosce):
    """Parole della grammatica che il modello non conosce (ordinate).
    ``conosce(parola) -> bool``. Vosk le ignorerebbe in silenzio (solo un
    avviso su stderr): vanno individuate prima, cosi' da togliere le frasi
    che le contengono e dire all'utente quali nomi non verranno capiti."""
    usate = {parola for frase in grammatica for parola in frase.split()}
    return sorted(p for p in usate if not conosce(p))


def filtra_grammatica(grammatica, sconosciute):
    """Grammatica senza le frasi che contengono parole sconosciute al modello."""
    escluse = set(sconosciute)
    return [frase for frase in grammatica if escluse.isdisjoint(frase.split())]


# ============================================================
# Assemblaggio
# ============================================================
def verifica_dipendenze():
    """Ritorna '' se tutto e' installato, altrimenti cosa manca (italiano)."""
    mancanti = [nome for nome, modulo in (("vosk", vosk), ("sounddevice", sounddevice),
                                          ("pyttsx3", pyttsx3)) if modulo is None]
    if mancanti:
        return ("Per la lettura vocale manca: " + ", ".join(mancanti) + ".\n"
                "Installa con:  pip install " + " ".join(mancanti)
                + "\n(su Linux può servire anche: sudo apt install libportaudio2 espeak-ng)")
    return ""


def crea_voce(grammatica, soglia=SOGLIA_CONFIDENZA_DEFAULT):
    """Crea (Ascoltatore, Sintesi, parole_sconosciute) pronti all'uso, gia'
    avviati. ``parole_sconosciute``: parole che il modello non conosce (le
    frasi che le contengono sono escluse dalla grammatica). Solleva VoceNonDisponibile con un messaggio chiaro se manca
    qualcosa (librerie, modello, microfono)."""
    manca = verifica_dipendenze()
    if manca:
        raise VoceNonDisponibile(manca)
    cartella = trova_modello()
    if cartella is None:
        raise VoceNonDisponibile(
            "Modello vocale italiano non trovato.\n"
            "Scarica 'vosk-model-small-it' da https://alphacephei.com/vosk/models "
            "e decomprimilo nella cartella 'modelli' accanto al programma.")

    modello = vosk.Model(str(cartella))
    sconosciute = parole_fuori_modello(
        grammatica, lambda parola: modello.vosk_model_find_word(parola) != -1)
    grammatica_ok = filtra_grammatica(grammatica, sconosciute)
    if not grammatica_ok:
        raise VoceNonDisponibile(
            "Il modello vocale non riconosce nessuna delle parole previste "
            "(controlla 'parola_attivazione' in pollencounter.cfg).")
    riconoscitore = vosk.KaldiRecognizer(modello, SAMPLE_RATE,
                                         json.dumps(grammatica_ok + ["[unk]"]))
    riconoscitore.SetWords(True)

    try:
        sorgente = SorgenteMicrofono()
    except Exception as e:
        raise VoceNonDisponibile(
            "Non riesco ad aprire il microfono. Controlla che sia collegato e "
            f"non usato da un altro programma.\n(dettaglio: {e})")

    ascoltatore = Ascoltatore(riconoscitore, sorgente, soglia)
    sintesi = Sintesi(ascoltatore)
    ascoltatore.avvia()
    return ascoltatore, sintesi, sconosciute
