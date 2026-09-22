import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import dominio
import sessione


LUNEDI = datetime(2026, 2, 9)  # e' effettivamente un lunedi'


class TestSettimana(unittest.TestCase):
    def setUp(self):
        self.s = sessione.Settimana(LUNEDI)

    def test_inserimento_singolo(self):
        self.s.attiva_giorno(1)
        totale = self.s.inserisci(1, "05", 1)
        self.assertEqual(totale, 1)
        self.assertEqual(self.s.conteggi["05"][0], 1)
        self.assertEqual(self.s.totale_giorno(1), 1)
        self.assertEqual(len(self.s.log), 1)
        self.assertEqual(self.s.log[0]["codice"], "05")

    def test_inserimento_multiplo_crea_piu_righe_di_log(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "48", 4)
        self.assertEqual(self.s.conteggi["48"][0], 4)
        self.assertEqual(len(self.s.log), 4)

    def test_annulla_ripristina_conteggio(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "05", 3)
        ris = self.s.annulla(1)
        self.assertEqual(ris, ("05", 3, 0))
        self.assertEqual(self.s.conteggi["05"][0], 0)

    def test_annulla_rimuove_le_righe_di_log_dell_inserimento(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "48", 4)
        n_log_prima = len(self.s.log)
        self.s.annulla(1)
        # le 4 righe dell'inserimento sono rimosse, ne resta 1 (ANNULLATO)
        self.assertEqual(len(self.s.log), n_log_prima - 4 + 1)
        self.assertTrue(any(r["nota"] and "ANNULLATO" in r["nota"] for r in self.s.log))

    def test_annulla_rimuove_la_voce_dallo_storico_difetto_7(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "05", 1)
        self.assertEqual(len(self.s.storico), 1)
        self.s.annulla(1)
        self.assertEqual(len(self.s.storico), 0)

    def test_annulla_senza_inserimenti_ritorna_none(self):
        self.s.attiva_giorno(1)
        self.assertIsNone(self.s.annulla(1))

    def test_annulla_fuori_dal_giorno_attivo_non_fa_nulla(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "05", 1)
        self.s.attiva_giorno(2)   # cambio giorno: azzera lo storico di annullo
        self.assertIsNone(self.s.annulla(1))
        self.assertEqual(self.s.conteggi["05"][0], 1)  # invariato

    def test_undo_non_intacca_note_intercalate(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "05", 1)
        self.s.aggiungi_nota(1, "nota di prova")
        self.s.annulla(1)
        # la nota deve sopravvivere, solo l'inserimento va rimosso
        note = [r for r in self.s.log if r["specie"] == "NOTA"]
        self.assertEqual(len(note), 1)
        self.assertEqual(note[0]["nota"], "nota di prova")

    def test_correggi(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "05", 2)
        vecchio = self.s.correggi(1, "05", 5)
        self.assertEqual(vecchio, 2)
        self.assertEqual(self.s.conteggi["05"][0], 5)

    def test_correggi_valore_negativo_rifiutato(self):
        with self.assertRaises(ValueError):
            self.s.correggi(1, "05", -1)

    def test_aggiungi_nota(self):
        self.s.aggiungi_nota(1, "prova")
        self.assertEqual(self.s.log[-1]["specie"], "NOTA")
        self.assertEqual(self.s.log[-1]["nota"], "prova")

    def test_ripeti_usa_ultimo_codice(self):
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "05", 1)
        self.assertEqual(self.s.ultimo_codice, "05")

    def test_listener_notificato_a_ogni_cambiamento(self):
        eventi = []
        self.s.on_change(lambda s: eventi.append(1))
        self.s.attiva_giorno(1)
        self.s.inserisci(1, "05", 1)
        self.assertEqual(len(eventi), 2)


class TestJournal(unittest.TestCase):
    def test_journal_scrive_snapshot_iniziale(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "~sessione_09-02-2026.jsonl"
            s = sessione.Settimana(LUNEDI)
            j = sessione.Journal(path)
            j.avvia(s)
            j.chiudi()
            righe = path.read_text(encoding="utf-8").strip().split("\n")
            self.assertEqual(len(righe), 1)
            primo = json.loads(righe[0])
            self.assertEqual(primo["t"], "inizio")
            self.assertEqual(primo["lunedi"], "2026-02-09")

    def test_replay_ricostruisce_stato_identico(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "~sessione_09-02-2026.jsonl"
            s = sessione.Settimana(LUNEDI)
            j = sessione.Journal(path)
            j.avvia(s)

            s.attiva_giorno(1, journal=j)
            s.inserisci(1, "05", 3, journal=j)
            s.inserisci(1, "48", 2, journal=j)
            s.annulla(1, journal=j)
            s.correggi(1, "05", 10, journal=j)
            s.aggiungi_nota(1, "controllato", journal=j)
            j.chiudi()

            ripristinata = sessione.ripristina_da_journal(path)
            self.assertEqual(ripristinata.lunedi, s.lunedi)
            self.assertEqual(ripristinata.fattore, s.fattore)
            self.assertEqual(ripristinata.conteggi, s.conteggi)
            self.assertEqual(len(ripristinata.log), len(s.log))

    def test_riprendi_non_riscrive_lo_snapshot(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "~sessione_09-02-2026.jsonl"
            s = sessione.Settimana(LUNEDI)
            j = sessione.Journal(path)
            j.avvia(s)
            s.attiva_giorno(1, journal=j)
            s.inserisci(1, "05", 1, journal=j)
            j.chiudi()

            j2 = sessione.Journal.riprendi(path)
            s2 = sessione.ripristina_da_journal(path)
            s2.inserisci(1, "48", 2, journal=j2)
            j2.chiudi()

            righe = path.read_text(encoding="utf-8").strip().split("\n")
            self.assertEqual(sum(1 for r in righe if json.loads(r)["t"] == "inizio"), 1)
            finale = sessione.ripristina_da_journal(path)
            self.assertEqual(finale.conteggi["05"][0], 1)
            self.assertEqual(finale.conteggi["48"][0], 2)

    def test_elimina_cancella_il_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "~sessione_09-02-2026.jsonl"
            s = sessione.Settimana(LUNEDI)
            j = sessione.Journal(path)
            j.avvia(s)
            j.elimina()
            self.assertFalse(path.exists())

    def test_recupera_sessioni(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            s = sessione.Settimana(LUNEDI)
            j = sessione.Journal(tmp / sessione.nome_journal(LUNEDI))
            j.avvia(s)
            s.attiva_giorno(1, journal=j)
            s.inserisci(1, "05", 1, journal=j)
            j.chiudi()

            risultati = sessione.recupera_sessioni(tmp)
            self.assertEqual(len(risultati), 1)
            path, info = risultati[0]
            self.assertEqual(info["lunedi"], "2026-02-09")
            self.assertEqual(info["n_operazioni"], 2)


class TestConfig(unittest.TestCase):
    def test_leggi_cartella_anno_assente(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "pollencounter.cfg"
            self.assertIsNone(sessione.leggi_cartella_anno(cfg, 2026))

    def test_salva_e_rileggi_cartella_anno(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            cfg = tmp / "pollencounter.cfg"
            cartella = tmp / "lavoro"
            cartella.mkdir()
            sessione.salva_cartella_anno(cfg, 2026, cartella)
            self.assertEqual(sessione.leggi_cartella_anno(cfg, 2026), cartella)

    def test_cartella_non_piu_esistente_ritorna_none(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            cfg = tmp / "pollencounter.cfg"
            cartella_fantasma = tmp / "sparita"
            sessione.salva_cartella_anno(cfg, 2026, cartella_fantasma)
            self.assertIsNone(sessione.leggi_cartella_anno(cfg, 2026))


if __name__ == "__main__":
    unittest.main()
