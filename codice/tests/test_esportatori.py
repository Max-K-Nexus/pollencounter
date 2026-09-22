import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl

import dominio
import esportatori
import sessione

LUNEDI = datetime(2026, 2, 9)


def _settimana_di_prova():
    s = sessione.Settimana(LUNEDI, fattore=0.4)
    s.attiva_giorno(1)
    s.inserisci(1, "01", 5)     # ACERACEAE
    s.inserisci(1, "04", 3)     # Alnus (fa parte di BETULACEAE)
    s.attiva_giorno(3)
    s.inserisci(3, "48", 10)    # Alternaria (spora)
    return s


class TestEsportaXlsx(unittest.TestCase):
    def test_scrive_conteggi_nelle_celle_giuste(self):
        s = _settimana_di_prova()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "prova.xlsx"
            esportatori.esporta_xlsx(s, path)
            wb = openpyxl.load_workbook(path)
            ws = wb["riepilogo_settimana"]
            # codice 01 -> riga 6, lunedi -> colonna G(7)
            self.assertEqual(ws.cell(row=6, column=7).value, 5)
            # codice 04 -> riga 9
            self.assertEqual(ws.cell(row=9, column=7).value, 3)
            # codice 48 -> riga 58, mercoledi -> colonna I(9)
            self.assertEqual(ws.cell(row=58, column=9).value, 10)
            self.assertEqual(ws["Q3"].value, 0.4)
            self.assertEqual(ws["M3"].value, 2026)
            wb.close()

    def test_scrive_il_log(self):
        s = _settimana_di_prova()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "prova.xlsx"
            esportatori.esporta_xlsx(s, path)
            wb = openpyxl.load_workbook(path)
            ws_log = wb["dati_grezzi"]
            n = sum(1 for row in ws_log.iter_rows(min_row=2, max_col=1) if row[0].value is not None)
            self.assertEqual(n, len(s.log))
            wb.close()

    def test_round_trip_import_export(self):
        """Esporta, poi ricarica con sessione.carica_da_xlsx: i conteggi devono
        tornare identici (verifica di parita' import/export)."""
        s = _settimana_di_prova()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "prova.xlsx"
            esportatori.esporta_xlsx(s, path)
            ripreso = sessione.carica_da_xlsx(path)
            self.assertEqual(ripreso.conteggi, s.conteggi)
            self.assertEqual(ripreso.fattore, s.fattore)


class TestCaricaSoglie(unittest.TestCase):
    def test_trova_file_nella_cartella_dello_script(self):
        soglie = esportatori.carica_soglie()
        self.assertIsNotNone(soglie)
        self.assertIn("Alternaria", soglie)
        # unica fonte: coerente col file, non col vecchio fallback nel codice
        self.assertEqual(soglie["Cladosporium"][0], 99.9)

    def test_none_se_non_trovato(self):
        with tempfile.TemporaryDirectory() as tmp:
            import percorsi
            vecchio_script_dir = percorsi.SCRIPT_DIR
            vecchio_bundle_dir = percorsi.BUNDLE_DIR
            try:
                percorsi.SCRIPT_DIR = Path(tmp)
                percorsi.BUNDLE_DIR = Path(tmp)
                esportatori.percorsi.SCRIPT_DIR = Path(tmp)
                esportatori.percorsi.BUNDLE_DIR = Path(tmp)
                self.assertIsNone(esportatori.carica_soglie(output_dir=tmp))
            finally:
                percorsi.SCRIPT_DIR = vecchio_script_dir
                percorsi.BUNDLE_DIR = vecchio_bundle_dir
                esportatori.percorsi.SCRIPT_DIR = vecchio_script_dir
                esportatori.percorsi.BUNDLE_DIR = vecchio_bundle_dir


class TestRiepilogoAnnuale(unittest.TestCase):
    def test_crea_file_con_foglio_settimanale_e_calendario(self):
        s = _settimana_di_prova()
        with tempfile.TemporaryDirectory() as tmp:
            percorso, n = esportatori.esporta_riepilogo_annuale(s, tmp)
            self.assertIsNotNone(percorso)
            self.assertEqual(n, 2)  # 2 giorni con dati (lunedi e mercoledi)
            wb = openpyxl.load_workbook(percorso)
            self.assertIn("Calendario", wb.sheetnames)
            self.assertIn("W07", wb.sheetnames)  # 9 feb 2026 e' nella settimana ISO 7
            wb.close()

    def test_nessun_dato_non_crea_file(self):
        s = sessione.Settimana(LUNEDI)
        with tempfile.TemporaryDirectory() as tmp:
            percorso, n = esportatori.esporta_riepilogo_annuale(s, tmp)
            self.assertIsNone(percorso)
            self.assertEqual(n, 0)

    def test_giorno_duplicato_somma_se_richiesto(self):
        s = _settimana_di_prova()
        with tempfile.TemporaryDirectory() as tmp:
            esportatori.esporta_riepilogo_annuale(s, tmp)
            # riesporta la stessa settimana: il lunedi' e' gia' presente
            esportatori.esporta_riepilogo_annuale(s, tmp, scegli_duplicati=lambda data: "c")
            percorso = Path(tmp) / "Riepilogo_Annuale_2026.xlsx"
            wb = openpyxl.load_workbook(percorso)
            ws = wb.active
            riga = esportatori.trova_riga_per_data(ws, "09/02/2026")
            col = esportatori._ann_col_grezzo("01")
            self.assertEqual(ws.cell(row=riga, column=col).value, 10)  # 5 + 5
            wb.close()

    def test_giorno_duplicato_annullato_non_modifica_il_dato_esistente(self):
        s = _settimana_di_prova()
        with tempfile.TemporaryDirectory() as tmp:
            esportatori.esporta_riepilogo_annuale(s, tmp)
            # riesporta la stessa settimana ma l'utente annulla: il dato gia'
            # presente non deve essere toccato e il giorno non va contato
            percorso, n = esportatori.esporta_riepilogo_annuale(
                s, tmp, scegli_duplicati=lambda data: "annulla")
            self.assertEqual(n, 0)
            wb = openpyxl.load_workbook(percorso)
            ws = wb.active
            riga = esportatori.trova_riga_per_data(ws, "09/02/2026")
            col = esportatori._ann_col_grezzo("01")
            self.assertEqual(ws.cell(row=riga, column=col).value, 5)  # invariato
            wb.close()


try:
    import docx  # noqa: F401
    _DOCX_DISPONIBILE = True
except ImportError:
    _DOCX_DISPONIBILE = False


@unittest.skipUnless(_DOCX_DISPONIBILE, "python-docx non installato")
class TestBollettiniWord(unittest.TestCase):
    def test_genera_ita_ed_eng(self):
        s = _settimana_di_prova()
        soglie = esportatori.carica_soglie() or {}
        with tempfile.TemporaryDirectory() as tmp:
            creati = esportatori.genera_bollettini_word(s, tmp, soglie=soglie)
            self.assertEqual(len(creati), 2)
            for p in creati:
                self.assertTrue(p.exists())

    def test_nessun_dato_non_genera_nulla(self):
        s = sessione.Settimana(LUNEDI)
        with tempfile.TemporaryDirectory() as tmp:
            creati = esportatori.genera_bollettini_word(s, tmp, soglie={})
            self.assertEqual(creati, [])


if __name__ == "__main__":
    unittest.main()
