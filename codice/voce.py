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
import math
import os
import queue
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import time
import wave
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

try:
    import winsound
except ImportError:
    winsound = None

SAMPLE_RATE = 16000
SOGLIA_CONFIDENZA_DEFAULT = 0.5
CODA_MUTO_SECONDI = 0.4          # coda di silenzio dopo la sintesi (eco nella stanza)
ATTESA_FRASE_DEFAULT = 0.3       # secondi di parziale invariato dopo i quali la frase si chiude
BLOCCO_MICROFONO = 800           # campioni per blocco audio (50 ms a 16 kHz)
NOME_MODELLO_PREFISSO = "vosk-model"


class VoceNonDisponibile(Exception):
    """Il messaggio e' scritto per l'utente (italiano, senza gergo)."""


# ============================================================
# Estrazione della frase da un risultato del riconoscitore (pura)
# ============================================================
def frase_da_risultato(risultato_json, soglia=SOGLIA_CONFIDENZA_DEFAULT):
    """Da un risultato Vosk (stringa JSON) ritorna la frase capita, oppure ''
    se vuota o se non c'e' nessuna parola sicura.

    Le parole sotto soglia di confidenza diventano '[unk]' invece di far
    scartare l'intera frase: con piu' comandi di seguito ("conta acero conta
    pioppo") un'incertezza deve rovinare solo il suo comando, non gli altri.
    Il comando che contiene '[unk]' viene poi scartato da
    ``dominio.interpreta_vocale`` (nel dubbio non si registra nulla: meglio
    ripetere che sbagliare specie)."""
    try:
        risultato = json.loads(risultato_json)
    except (TypeError, ValueError):
        return ""
    testo = (risultato.get("text") or "").strip()
    if not testo:
        return ""
    parole = risultato.get("result") or []
    if parole:
        testo = " ".join(p.get("word", "") if p.get("conf", 1.0) >= soglia else "[unk]"
                         for p in parole)
    if all(p == "[unk]" for p in testo.split()):
        return ""
    return testo


# ============================================================
# Ascolto
# ============================================================
class Ascoltatore:
    """Thread di ascolto: blocchi audio -> riconoscitore -> coda di frasi.

    ``sorgente_audio``: iterabile (bloccante) di blocchi ``bytes`` (16 bit,
    mono, ``SAMPLE_RATE``).
    ``riconoscitore``: oggetto con AcceptWaveform(bytes)->bool, Result()->str
    (JSON) e Reset(); per la chiusura anticipata anche PartialResult() e
    FinalResult().

    Chiusura anticipata: Vosk dichiara finita una frase solo dopo circa un
    secondo di silenzio, troppo per chi conta pollini a ritmo serrato. Se
    ``frase_completa`` e' dato, la frase viene chiusa non appena il risultato
    parziale e' un comando valido (``frase_completa(parziale) -> bool``) e non
    cambia da ``attesa`` secondi di audio. Il tempo e' misurato sull'audio
    ricevuto, non sull'orologio. La chiusura normale di Vosk resta attiva.

    Mentre e' "muto" (sintesi vocale in corso, se si usa l'anti-eco) l'audio
    viene scartato e il riconoscitore azzerato: altrimenti il PC sentirebbe se
    stesso dire "graminacee" e registrerebbe un secondo granulo.
    """

    def __init__(self, riconoscitore, sorgente_audio, soglia=SOGLIA_CONFIDENZA_DEFAULT,
                 frase_completa=None, attesa=ATTESA_FRASE_DEFAULT):
        self._riconoscitore = riconoscitore
        self._sorgente = sorgente_audio
        self._soglia = soglia
        self._frase_completa = frase_completa
        self._attesa = attesa
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

    def _metti_in_coda(self, risultato_json):
        frase = frase_da_risultato(risultato_json, self._soglia)
        if frase:
            self._frasi.put(frase)

    def _parziale(self):
        try:
            return (json.loads(self._riconoscitore.PartialResult()).get("partial") or "").strip()
        except (TypeError, ValueError):
            return ""

    def _ciclo(self):
        campioni = 0                # audio ricevuto finora (in campioni: niente arrotondamenti)
        attesa_campioni = round(self._attesa * SAMPLE_RATE)
        parziale, dal = "", 0       # ultimo parziale e campione in cui e' apparso
        try:
            for blocco in self._sorgente:
                if self._ferma.is_set():
                    break
                if time.monotonic() < self._muto_fino:
                    self._riconoscitore.Reset()
                    parziale = ""
                    continue
                campioni += len(blocco) // 2
                if self._riconoscitore.AcceptWaveform(blocco):
                    parziale = ""
                    self._metti_in_coda(self._riconoscitore.Result())
                elif self._frase_completa is not None:
                    attuale = self._parziale()
                    if attuale != parziale:
                        parziale, dal = attuale, campioni
                    elif (attuale and campioni - dal >= attesa_campioni
                          and self._frase_completa(attuale)):
                        parziale = ""
                        self._metti_in_coda(self._riconoscitore.FinalResult())
        except Exception as e:          # microfono scollegato, driver audio, ecc.
            self.errore = str(e)


class SorgenteMicrofono:
    """Blocchi audio dal microfono predefinito (sounddevice)."""

    def __init__(self):
        self._blocchi = queue.Queue()
        self._chiuso = False
        self._stream = sounddevice.RawInputStream(
            samplerate=SAMPLE_RATE, blocksize=BLOCCO_MICROFONO, dtype="int16", channels=1,
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
# Riproduzione audio (beep e talkback)
#
# Su Linux non ci si puo' affidare al "campanello" del desktop (di solito
# disattivato) ne' a pyttsx3, che riproduce il suono chiamando `aplay`
# (pacchetto alsa-utils, spesso assente con PipeWire/PulseAudio). Si usa
# invece il primo lettore di file WAV che c'e' sul sistema.
# ============================================================
BEEP_HZ = 800
BEEP_MS = 150
_beep_wav = None                    # percorso del WAV del beep, creato al primo uso


def trova_lettore_audio(which=shutil.which, piattaforma=None):
    """Comando (lista) per riprodurre un file WAV passandogli il percorso in
    coda, o None se non c'e' (su Windows si usa winsound, quindi None)."""
    piattaforma = piattaforma or sys.platform
    if piattaforma.startswith("win"):
        return None
    if piattaforma == "darwin":
        candidati = [["afplay"]]
    else:
        candidati = [["paplay"], ["pw-play"], ["aplay", "-q"]]
    for comando in candidati:
        percorso = which(comando[0])
        if percorso:
            return [percorso] + comando[1:]
    return None


def scrivi_wav_beep(percorso, hz=BEEP_HZ, ms=BEEP_MS, rate=22050):
    """Scrive un breve tono (mono, 16 bit) con attacco/rilascio morbidi, per
    evitare lo 'scatto' all'inizio e alla fine."""
    n = int(rate * ms / 1000)
    rampa = max(1, int(rate * 0.01))
    campioni = []
    for i in range(n):
        ampiezza = min(1.0, i / rampa, (n - 1 - i) / rampa)
        campioni.append(int(12000 * ampiezza * math.sin(2 * math.pi * hz * i / rate)))
    with wave.open(str(percorso), "wb") as f:
        f.setnchannels(1)
        f.setsampwidth(2)
        f.setframerate(rate)
        f.writeframes(struct.pack(f"<{n}h", *campioni))


def riproduci_wav(percorso, attendi=True, lettore=None):
    """Riproduce un WAV. Ritorna True se la riproduzione e' partita (e, con
    ``attendi``, terminata senza errori), False altrimenti: mai eccezioni."""
    if sys.platform.startswith("win"):
        if winsound is None:
            return False
        try:
            flag = winsound.SND_FILENAME | (0 if attendi else winsound.SND_ASYNC)
            winsound.PlaySound(str(percorso), flag)
            return True
        except Exception:
            return False
    comando = lettore or trova_lettore_audio()
    if not comando:
        return False
    try:
        if attendi:
            subprocess.run(comando + [str(percorso)], check=True, timeout=30,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            processo = subprocess.Popen(comando + [str(percorso)],
                                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            threading.Thread(target=processo.wait, daemon=True).start()    # niente zombie
        return True
    except Exception:
        return False


def beep():
    """Beep di conferma udibile su Windows, Linux e macOS. Ritorna True se e'
    partito un suono. Non blocca (tranne i 150 ms di winsound.Beep)."""
    global _beep_wav
    if sys.platform.startswith("win"):
        if winsound is None:
            return False
        try:
            winsound.Beep(BEEP_HZ, BEEP_MS)
            return True
        except Exception:
            return False
    try:
        if _beep_wav is None or not os.path.exists(_beep_wav):
            fd, percorso = tempfile.mkstemp(prefix="polline_beep_", suffix=".wav")
            os.close(fd)
            scrivi_wav_beep(percorso)
            _beep_wav = percorso
    except Exception:
        return False
    return riproduci_wav(_beep_wav, attendi=False)


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


class MotorePyttsx3:
    """Sintesi con pyttsx3 (SAPI5 su Windows, NSSpeech su macOS). Va creato
    nel thread che lo usera' (richiesto da COM su Windows)."""

    def __init__(self):
        self._motore = pyttsx3.init()
        self.avviso = ""
        id_voce = scegli_voce_italiana(self._motore.getProperty("voices"))
        if id_voce:
            self._motore.setProperty("voice", id_voce)
        else:
            self.avviso = ("Nessuna voce italiana installata: il computer pronuncera' "
                           "i nomi con una voce straniera.")

    def dici(self, testo):
        self._motore.say(testo)
        self._motore.runAndWait()


class MotoreEspeak:
    """Sintesi con espeak-ng da riga di comando + lettore WAV del sistema
    (paplay/pw-play/aplay). Evita pyttsx3 su Linux, che richiede `aplay`."""

    def __init__(self, eseguibile, lettore):
        self._eseguibile = eseguibile
        self._lettore = lettore
        self.avviso = ""

    def dici(self, testo):
        fd, percorso = tempfile.mkstemp(prefix="polline_voce_", suffix=".wav")
        os.close(fd)
        try:
            subprocess.run([self._eseguibile, "-v", "it", "-w", percorso, "--", testo],
                           check=True, timeout=15,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if not riproduci_wav(percorso, attendi=True, lettore=self._lettore):
                raise RuntimeError("non riesco a riprodurre l'audio")
        finally:
            try:
                os.remove(percorso)
            except OSError:
                pass


def crea_motore_sintesi(which=shutil.which, piattaforma=None):
    """Sceglie il motore di sintesi: espeak-ng + lettore WAV su Linux se ci
    sono entrambi, altrimenti pyttsx3. Ritorna (motore, avviso)."""
    piattaforma = piattaforma or sys.platform
    avviso = ""
    if piattaforma.startswith("linux"):
        lettore = trova_lettore_audio(which, piattaforma)
        eseguibile = which("espeak-ng") or which("espeak")
        if lettore and eseguibile:
            return MotoreEspeak(eseguibile, lettore), ""
        if not lettore:
            avviso = ("Non trovo un programma per riprodurre l'audio: la voce restera' muta. "
                      "Installa 'pulseaudio-utils' o 'alsa-utils' "
                      "(es. sudo zypper install alsa-utils).")
    motore = MotorePyttsx3()
    return motore, avviso or motore.avviso


class Sintesi:
    """Il PC dice a voce cio' che ha capito. Un solo thread possiede il
    motore di sintesi (richiesto da pyttsx3/COM su Windows); le frasi
    arrivano da una coda. Se c'e' un ascoltatore (anti-eco), prima di parlare
    lo silenzia, e lo silenzia di nuovo a fine frase (la coda di eco parte da
    quando l'audio e' davvero finito). Se le conferme si accumulano (si conta
    piu' in fretta di quanto il PC riesca a parlare) vengono dette solo le
    piu' recenti: la voce non resta indietro di secondi.

    ``motore``: oggetto con ``dici(testo)`` (bloccante); se None lo sceglie
    ``crea_motore_sintesi`` nel thread di sintesi.
    ``avviso``: se non vuoto, la sintesi non e' pienamente disponibile
    (nessuna voce italiana, nessun lettore audio, motore non avviabile):
    da mostrare all'utente."""

    def __init__(self, ascoltatore=None, durata_stimata=lambda testo: 0.5 + 0.08 * len(testo),
                 motore=None):
        self._ascoltatore = ascoltatore
        self._durata = durata_stimata
        self._motore = motore
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
        motore = self._motore
        try:
            if motore is None:
                motore, self.avviso = crea_motore_sintesi()
        except Exception as e:
            motore = None
            self.avviso = f"Sintesi vocale non disponibile ({e}): niente conferma parlata."
        finally:
            self._pronta.set()
        while True:
            testo = self._coda.get()
            ultima = testo is None
            while not ultima:
                try:
                    altro = self._coda.get_nowait()
                except queue.Empty:
                    break
                if altro is None:
                    ultima = True               # stop richiesto: dopo questa frase
                else:
                    testo = altro               # salta la conferma arretrata
            if testo is None:
                break
            if motore is None:
                if ultima:
                    break
                continue
            try:
                if self._ascoltatore is not None:
                    self._ascoltatore.silenzia(self._durata(testo))
                motore.dici(testo)
                if self._ascoltatore is not None:
                    self._ascoltatore.silenzia(0)
            except Exception as e:      # un errore di sintesi non deve fermare la conta
                self.avviso = f"Errore della sintesi vocale: {e}"
            if ultima:
                break


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


def crea_voce(grammatica, soglia=SOGLIA_CONFIDENZA_DEFAULT, frase_completa=None,
              attesa=ATTESA_FRASE_DEFAULT, anti_eco=True):
    """Crea (Ascoltatore, Sintesi, parole_sconosciute) pronti all'uso, gia'
    avviati. ``parole_sconosciute``: parole che il modello non conosce (le
    frasi che le contengono sono escluse dalla grammatica).
    ``frase_completa``/``attesa``: chiusura anticipata della frase (vedi
    ``Ascoltatore``); ``attesa`` <= 0 la disattiva.
    ``anti_eco``: se False il microfono resta aperto mentre il PC parla (va
    bene solo se la voce del PC non puo' essere scambiata per un comando,
    cioe' con la parola di attivazione). Solleva VoceNonDisponibile con un
    messaggio chiaro se manca qualcosa (librerie, modello, microfono)."""
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

    ascoltatore = Ascoltatore(riconoscitore, sorgente, soglia,
                              frase_completa if attesa > 0 else None, attesa)
    sintesi = Sintesi(ascoltatore if anti_eco else None)
    ascoltatore.avvia()
    return ascoltatore, sintesi, sconosciute
