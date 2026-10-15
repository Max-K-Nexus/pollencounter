import json
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest import mock

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

    def test_imposta_fattore_notifica_e_ritorna_il_vecchio(self):
        s = sessione.Settimana(LUNEDI)
        chiamate = []
        s.on_change(chiamate.append)
        self.assertEqual(s.imposta_fattore(0.5), 0.4)
        self.assertEqual(s.fattore, 0.5)
        self.assertEqual(len(chiamate), 1)
        with self.assertRaises(ValueError):
            s.imposta_fattore(0)
        self.assertEqual(s.fattore, 0.5)

    def test_replay_ripristina_il_fattore_cambiato(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "~sessione_09-02-2026.jsonl"
            s = sessione.Settimana(LUNEDI)
            j = sessione.Journal(path)
            j.avvia(s)
            s.attiva_giorno(1, journal=j)
            s.inserisci(1, "05", 3, journal=j)
            s.imposta_fattore(0.55, journal=j)
            j.chiudi()
            self.assertEqual(sessione.ripristina_da_journal(path).fattore, 0.55)

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

    def test_rimuovi_sessioni(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            altro_lunedi = LUNEDI + timedelta(days=7)
            for lunedi in (LUNEDI, altro_lunedi):
                j = sessione.Journal(tmp / sessione.nome_journal(lunedi))
                j.avvia(sessione.Settimana(lunedi))
                j.chiudi()
            self.assertEqual(len(sessione.recupera_sessioni(tmp)), 2)

            # un file gia' sparito non e' un errore
            falliti = sessione.rimuovi_sessioni(
                [tmp / sessione.nome_journal(LUNEDI), tmp / "~sessione_inesistente.jsonl"])
            self.assertEqual(falliti, [])
            rimaste = sessione.recupera_sessioni(tmp)
            self.assertEqual(len(rimaste), 1)
            self.assertEqual(rimaste[0][0].name, sessione.nome_journal(altro_lunedi))

            self.assertEqual(sessione.rimuovi_sessioni([]), [])


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

    def _cfg(self, tmp, contenuto):
        cfg = Path(tmp) / "pollencounter.cfg"
        cfg.write_text(json.dumps(contenuto), encoding="utf-8")
        return cfg

    def test_layout_finestra_assente_o_malformato(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(sessione.leggi_layout_finestra(Path(tmp) / "nonesiste.cfg"), {})
            self.assertEqual(sessione.leggi_layout_finestra(self._cfg(tmp, {"finestra": "x"})), {})

    def test_layout_finestra_scarta_solo_i_valori_non_validi(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = self._cfg(tmp, {"finestra": {"larghezza": 1100, "altezza": "alta",
                                               "divisione": 0.95}})
            self.assertEqual(sessione.leggi_layout_finestra(cfg), {"larghezza": 1100})

    def test_layout_finestra_salvato_e_riletto_senza_toccare_il_resto(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = self._cfg(tmp, {"2026": "/da/qualche/parte", "voce_attesa_ms": 250})
            sessione.salva_layout_finestra(cfg, 1280, 720, 0.6)
            self.assertEqual(sessione.leggi_layout_finestra(cfg),
                             {"larghezza": 1280, "altezza": 720, "divisione": 0.6})
            config = json.loads(cfg.read_text(encoding="utf-8"))
            self.assertEqual(config["2026"], "/da/qualche/parte")
            self.assertEqual(config["voce_attesa_ms"], 250)

    def test_layout_finestra_valori_assurdi_non_vengono_salvati(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = Path(tmp) / "pollencounter.cfg"
            sessione.salva_layout_finestra(cfg, 1, 1, 0.01)      # finestra ridotta a icona
            self.assertFalse(cfg.exists())

    def test_sinonimi_vocali_assenti(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(sessione.leggi_sinonimi_vocali(Path(tmp) / "nonesiste.cfg"), {})
            self.assertEqual(sessione.leggi_sinonimi_vocali(self._cfg(tmp, {"2026": "x"})), {})

    def test_sinonimi_vocali_letti_e_ripuliti(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = self._cfg(tmp, {"sinonimi_vocali": {"24": ["erba", 3], "05": "no"}})
            self.assertEqual(sessione.leggi_sinonimi_vocali(cfg), {"24": ["erba"]})

    def test_sinonimi_vocali_sopravvivono_al_salvataggio_cartella(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = self._cfg(tmp, {"sinonimi_vocali": {"24": ["erba"]}})
            sessione.salva_cartella_anno(cfg, 2026, Path(tmp))
            self.assertEqual(sessione.leggi_sinonimi_vocali(cfg), {"24": ["erba"]})

    def test_parola_attivazione_default(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(sessione.leggi_parola_attivazione(Path(tmp) / "nonesiste.cfg"), "conta")
            self.assertEqual(sessione.leggi_parola_attivazione(self._cfg(tmp, {"parola_attivazione": 5})), "conta")

    def test_parola_attivazione_personalizzata_e_disattivata(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(sessione.leggi_parola_attivazione(
                self._cfg(tmp, {"parola_attivazione": "Registra"})), "registra")
            self.assertEqual(sessione.leggi_parola_attivazione(
                self._cfg(tmp, {"parola_attivazione": ""})), "")

    def test_attesa_vocale(self):
        with tempfile.TemporaryDirectory() as tmp:
            leggi = sessione.leggi_attesa_vocale
            self.assertEqual(leggi(Path(tmp) / "nonesiste.cfg"), 300)
            self.assertEqual(leggi(self._cfg(tmp, {"voce_attesa_ms": 450})), 450)
            self.assertEqual(leggi(self._cfg(tmp, {"voce_attesa_ms": 0})), 0)      # disattivata
            self.assertEqual(leggi(self._cfg(tmp, {"voce_attesa_ms": 99999})), 2000)
            for non_valido in ("tanto", -5, None, True):
                self.assertEqual(leggi(self._cfg(tmp, {"voce_attesa_ms": non_valido})), 300)


@unittest.skipIf(sessione.openpyxl is None, "openpyxl non installato")
class TestContaRigheLog(unittest.TestCase):
    def setUp(self):
        sessione._conta_righe_log.cache_clear()

    def _xlsx_con_log(self, tmp, righe):
        wb = sessione.openpyxl.Workbook()
        ws = wb.active
        ws.title = "dati_grezzi"
        ws.append(["data"])
        for i in range(righe):
            ws.append([f"riga {i}"])
        percorso = Path(tmp) / "conta.xlsx"
        wb.save(percorso)
        return percorso

    def test_conta_le_righe_dopo_l_intestazione(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(sessione.conta_righe_log(self._xlsx_con_log(tmp, 3)), "(3 righe log)")

    def test_errore_transitorio_non_resta_in_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            percorso = self._xlsx_con_log(tmp, 2)
            vero = sessione.openpyxl.load_workbook
            with mock.patch.object(sessione.openpyxl, "load_workbook",
                                   side_effect=PermissionError("file bloccato")):
                self.assertEqual(sessione.conta_righe_log(percorso), "")
            # Stesso file, stessa data di modifica: ora e' leggibile e va letto davvero.
            with mock.patch.object(sessione.openpyxl, "load_workbook", side_effect=vero):
                self.assertEqual(sessione.conta_righe_log(percorso), "(2 righe log)")

    def test_file_inesistente(self):
        self.assertEqual(sessione.conta_righe_log("/non/esiste.xlsx"), "")


if __name__ == "__main__":
    unittest.main()
