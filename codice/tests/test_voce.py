"""Test di voce.py con un riconoscitore finto: non servono microfono, Vosk
ne' sintesi vocale (l'integrazione col motore reale si prova a mano, vedi
CHANGELOG)."""

import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

import voce


def risultato(testo, conf=1.0):
    return json.dumps({"text": testo,
                       "result": [{"word": w, "conf": conf} for w in testo.split()]})


class RiconoscitoreFinto:
    """Un blocco audio = un risultato pronto (None = frase ancora in corso)."""

    def __init__(self, risultati):
        self._risultati = list(risultati)
        self._ultimo = None
        self.reset = 0

    def AcceptWaveform(self, _blocco):
        self._ultimo = self._risultati.pop(0)
        return self._ultimo is not None

    def Result(self):
        return self._ultimo

    def Reset(self):
        self.reset += 1


def esegui(risultati, muto_secondi=0.0, soglia=0.5):
    """Fa girare un Ascoltatore su tanti blocchi quanti sono i risultati e
    ritorna (frasi capite, riconoscitore)."""
    riconoscitore = RiconoscitoreFinto(risultati)
    asc = voce.Ascoltatore(riconoscitore, [b"x"] * len(risultati), soglia)
    if muto_secondi:
        asc.silenzia(muto_secondi)
        asc._muto_fino = time.monotonic() + 60       # muto per tutta la prova
    asc.avvia()
    asc._thread.join(timeout=2)
    frasi = []
    while (f := asc.prossima_frase()) is not None:
        frasi.append(f)
    return frasi, riconoscitore, asc


class TestFraseDaRisultato(unittest.TestCase):
    def test_frase_capita(self):
        self.assertEqual(voce.frase_da_risultato(risultato("conta ventiquattro")),
                         "conta ventiquattro")

    def test_vuota_o_malformata(self):
        for r in ("", "non json", None, json.dumps({}), json.dumps({"text": "  "})):
            self.assertEqual(voce.frase_da_risultato(r), "")

    def test_unk_scartato(self):
        self.assertEqual(voce.frase_da_risultato(risultato("conta [unk]")), "")

    def test_confidenza_sotto_soglia_scartata(self):
        self.assertEqual(voce.frase_da_risultato(risultato("conta ventiquattro", conf=0.3)), "")
        self.assertEqual(voce.frase_da_risultato(risultato("conta ventiquattro", conf=0.3), soglia=0.2),
                         "conta ventiquattro")

    def test_basta_una_parola_incerta(self):
        r = json.dumps({"text": "conta tre", "result": [{"word": "conta", "conf": 1.0},
                                                         {"word": "tre", "conf": 0.2}]})
        self.assertEqual(voce.frase_da_risultato(r), "")


class TestAscoltatore(unittest.TestCase):
    def test_frasi_in_coda_nell_ordine(self):
        frasi, _, _ = esegui([risultato("conta uno"), None, risultato("conta annulla")])
        self.assertEqual(frasi, ["conta uno", "conta annulla"])

    def test_frasi_incerte_o_sconosciute_non_arrivano(self):
        frasi, _, _ = esegui([risultato("conta [unk]"), risultato("conta tre", conf=0.1),
                              risultato("conta tre")])
        self.assertEqual(frasi, ["conta tre"])

    def test_nessuna_frase_ritorna_none(self):
        asc = voce.Ascoltatore(RiconoscitoreFinto([]), [], 0.5)
        self.assertIsNone(asc.prossima_frase())

    def test_audio_scartato_mentre_muto(self):
        # anti-eco: il PC non deve registrare la propria voce
        frasi, riconoscitore, _ = esegui([risultato("conta graminacee")] * 3, muto_secondi=1)
        self.assertEqual(frasi, [])
        self.assertEqual(riconoscitore.reset, 3)

    def test_silenzia_non_accorcia_un_silenzio_piu_lungo(self):
        asc = voce.Ascoltatore(RiconoscitoreFinto([]), [], 0.5)
        asc.silenzia(10)
        lungo = asc._muto_fino
        asc.silenzia(0.1)
        self.assertEqual(asc._muto_fino, lungo)

    def test_errore_del_microfono_non_blocca_ne_manda_in_crash(self):
        def sorgente():
            yield b"x"
            raise OSError("microfono scollegato")

        asc = voce.Ascoltatore(RiconoscitoreFinto([risultato("conta uno")]), sorgente(), 0.5)
        asc.avvia()
        asc._thread.join(timeout=2)
        self.assertEqual(asc.prossima_frase(), "conta uno")
        self.assertIn("microfono scollegato", asc.errore)


class TestVocabolarioModello(unittest.TestCase):
    CONOSCIUTE = {"conta", "ventiquattro", "per", "tre", "ortica"}

    def test_parole_fuori_modello(self):
        grammatica = ["conta ventiquattro", "conta stemphylium per tre"]
        self.assertEqual(voce.parole_fuori_modello(grammatica, self.CONOSCIUTE.__contains__),
                         ["stemphylium"])

    def test_nessuna_parola_fuori_modello(self):
        self.assertEqual(voce.parole_fuori_modello(["conta ortica"], self.CONOSCIUTE.__contains__), [])

    def test_filtra_grammatica_toglie_le_frasi_con_parole_ignote(self):
        grammatica = ["conta ventiquattro", "conta urticaceae", "conta urticaceae per tre", "conta ortica"]
        self.assertEqual(voce.filtra_grammatica(grammatica, ["urticaceae"]),
                         ["conta ventiquattro", "conta ortica"])

    def test_filtra_grammatica_senza_sconosciute_non_cambia(self):
        grammatica = ["conta uno", "conta due"]
        self.assertEqual(voce.filtra_grammatica(grammatica, []), grammatica)

    def test_trova_modello(self):
        with tempfile.TemporaryDirectory() as tmp:
            modello = Path(tmp) / "modelli" / "vosk-model-small-it-0.22"
            modello.mkdir(parents=True)
            (Path(tmp) / "modelli" / "altro").mkdir()
            with mock.patch.object(voce.percorsi, "BUNDLE_DIR", Path(tmp)):
                self.assertEqual(voce.trova_modello(), modello)

    def test_trova_modello_assente(self):
        with tempfile.TemporaryDirectory() as tmp, \
                mock.patch.object(voce.percorsi, "BUNDLE_DIR", Path(tmp)), \
                mock.patch.object(voce.percorsi, "SCRIPT_DIR", Path(tmp)), \
                mock.patch.object(voce.percorsi, "EXE_DIR", Path(tmp)):
            self.assertIsNone(voce.trova_modello())


class VoceFinta:
    def __init__(self, id, name="", languages=()):
        self.id, self.name, self.languages = id, name, list(languages)


class TestSceltaVoceItaliana(unittest.TestCase):
    def test_espeak_per_id_e_lingua(self):
        voci = [VoceFinta("gmw/en", "English", ["en"]), VoceFinta("sit/cmn", "Chinese", ["cmn"]),
                VoceFinta("roa/it", "Italian", ["it"])]
        self.assertEqual(voce.scegli_voce_italiana(voci), "roa/it")

    def test_lingua_in_bytes_con_byte_di_lunghezza(self):
        voci = [VoceFinta("x1", "Voce uno", [b"\x02it"])]
        self.assertEqual(voce.scegli_voce_italiana(voci), "x1")

    def test_windows_sapi_per_nome(self):
        voci = [VoceFinta("HKEY\\EN-US", "Microsoft David Desktop - English (United States)"),
                VoceFinta("HKEY\\IT-IT", "Microsoft Elsa Desktop - Italian (Italy)")]
        self.assertEqual(voce.scegli_voce_italiana(voci), "HKEY\\IT-IT")

    def test_it_dentro_un_altro_id_non_basta(self):
        # 'sit/cmn' contiene 'it' ma e' cinese: errore reale preso in prova
        self.assertIsNone(voce.scegli_voce_italiana([VoceFinta("sit/cmn", "Chinese", ["cmn"])]))

    def test_nessuna_voce_italiana(self):
        self.assertIsNone(voce.scegli_voce_italiana([VoceFinta("gmw/en", "English", ["en"])]))
        self.assertIsNone(voce.scegli_voce_italiana([]))


class TestDipendenze(unittest.TestCase):
    def test_messaggio_se_mancano_librerie(self):
        with mock.patch.object(voce, "vosk", None), mock.patch.object(voce, "pyttsx3", None):
            msg = voce.verifica_dipendenze()
        self.assertIn("vosk", msg)
        self.assertIn("pyttsx3", msg)
        self.assertIn("pip install", msg)

    def test_crea_voce_senza_librerie_spiega_cosa_manca(self):
        with mock.patch.object(voce, "vosk", None):
            with self.assertRaises(voce.VoceNonDisponibile) as ctx:
                voce.crea_voce(["conta uno"])
        self.assertIn("vosk", str(ctx.exception))

    def test_crea_voce_senza_modello_spiega_dove_metterlo(self):
        finto = mock.Mock()
        with mock.patch.object(voce, "vosk", finto), mock.patch.object(voce, "sounddevice", finto), \
                mock.patch.object(voce, "pyttsx3", finto), \
                mock.patch.object(voce, "trova_modello", return_value=None):
            with self.assertRaises(voce.VoceNonDisponibile) as ctx:
                voce.crea_voce(["conta uno"])
        self.assertIn("modelli", str(ctx.exception))


def which_finto(*presenti):
    return lambda nome: f"/usr/bin/{nome}" if nome in presenti else None


class TestLettoreAudio(unittest.TestCase):
    def test_ordine_di_preferenza_su_linux(self):
        self.assertEqual(voce.trova_lettore_audio(which_finto("aplay", "pw-play", "paplay"), "linux"),
                         ["/usr/bin/paplay"])
        self.assertEqual(voce.trova_lettore_audio(which_finto("aplay", "pw-play"), "linux"),
                         ["/usr/bin/pw-play"])
        self.assertEqual(voce.trova_lettore_audio(which_finto("aplay"), "linux"),
                         ["/usr/bin/aplay", "-q"])

    def test_nessun_lettore(self):
        self.assertIsNone(voce.trova_lettore_audio(which_finto(), "linux"))

    def test_windows_usa_winsound_quindi_nessun_comando(self):
        self.assertIsNone(voce.trova_lettore_audio(which_finto("paplay"), "win32"))

    def test_macos(self):
        self.assertEqual(voce.trova_lettore_audio(which_finto("afplay"), "darwin"),
                         ["/usr/bin/afplay"])

    def test_wav_del_beep_valido(self):
        import wave
        with tempfile.TemporaryDirectory() as d:
            percorso = Path(d) / "beep.wav"
            voce.scrivi_wav_beep(percorso)
            with wave.open(str(percorso), "rb") as f:
                self.assertEqual(f.getnchannels(), 1)
                self.assertEqual(f.getsampwidth(), 2)
                durata_ms = 1000 * f.getnframes() / f.getframerate()
        self.assertAlmostEqual(durata_ms, voce.BEEP_MS, delta=5)

    def test_riproduci_senza_lettore_non_solleva(self):
        with mock.patch.object(voce.sys, "platform", "linux"), \
                mock.patch.object(voce, "trova_lettore_audio", return_value=None):
            self.assertFalse(voce.riproduci_wav("x.wav"))

    def test_beep_senza_lettore_ritorna_false(self):
        with mock.patch.object(voce.sys, "platform", "linux"), \
                mock.patch.object(voce, "trova_lettore_audio", return_value=None):
            self.assertFalse(voce.beep())


class MotoreFinto:
    def __init__(self, errore_alla=None):
        self.detto = []
        self._errore_alla = errore_alla

    def dici(self, testo):
        if testo == self._errore_alla:
            raise RuntimeError("audio guasto")
        self.detto.append(testo)


class AscoltatoreFinto:
    def __init__(self):
        self.chiamate = []

    def silenzia(self, secondi):
        self.chiamate.append(secondi)


class TestSintesi(unittest.TestCase):
    def _attendi(self, sintesi):
        sintesi.ferma()
        sintesi._thread.join(timeout=2)

    def test_il_testo_arriva_al_motore_e_l_ascoltatore_e_silenziato_prima_e_dopo(self):
        motore, asc = MotoreFinto(), AscoltatoreFinto()
        sintesi = voce.Sintesi(asc, durata_stimata=lambda t: 1.5, motore=motore)
        sintesi.parla("graminacee")
        self._attendi(sintesi)
        self.assertEqual(motore.detto, ["graminacee"])
        self.assertEqual(asc.chiamate, [1.5, 0])

    def test_errore_del_motore_non_ferma_il_thread(self):
        motore = MotoreFinto(errore_alla="uno")
        sintesi = voce.Sintesi(None, motore=motore)
        sintesi.parla("uno")
        sintesi.parla("due")
        self._attendi(sintesi)
        self.assertEqual(motore.detto, ["due"])
        self.assertIn("audio guasto", sintesi.avviso)

    def test_motore_non_avviabile_da_un_avviso(self):
        with mock.patch.object(voce, "crea_motore_sintesi", side_effect=RuntimeError("no driver")):
            sintesi = voce.Sintesi(None)
            sintesi.parla("uno")
            self._attendi(sintesi)
        self.assertIn("no driver", sintesi.avviso)


class TestSceltaMotoreSintesi(unittest.TestCase):
    def test_linux_con_espeak_e_lettore_non_usa_pyttsx3(self):
        motore, avviso = voce.crea_motore_sintesi(which_finto("espeak-ng", "paplay"), "linux")
        self.assertIsInstance(motore, voce.MotoreEspeak)
        self.assertEqual(avviso, "")

    def test_linux_senza_lettore_avvisa_e_dice_cosa_installare(self):
        finto = mock.Mock()
        finto.init.return_value.getProperty.return_value = []
        with mock.patch.object(voce, "pyttsx3", finto):
            motore, avviso = voce.crea_motore_sintesi(which_finto("espeak-ng"), "linux")
        self.assertIsInstance(motore, voce.MotorePyttsx3)
        self.assertIn("alsa-utils", avviso)


if __name__ == "__main__":
    unittest.main()
