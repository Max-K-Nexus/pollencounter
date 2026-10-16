import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

import archivio


LUNEDI = datetime(2026, 10, 12)      # settimana ISO 42
ALTRO_LUNEDI = datetime(2026, 10, 5)  # settimana ISO 41


def _crea(path):
    """Crea un file vuoto (e le cartelle che servono)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.touch()
    return path


class TestNomi(unittest.TestCase):
    def test_nome_cartella_settimana(self):
        self.assertEqual(archivio.nome_cartella_settimana(LUNEDI), "S42_2026-10-12")
        self.assertEqual(archivio.nome_cartella_settimana(ALTRO_LUNEDI), "S41_2026-10-05")

    def test_numero_settimana_sempre_a_due_cifre(self):
        # in ordine alfabetico = in ordine cronologico, anche per le prime settimane
        nomi = [archivio.nome_cartella_settimana(datetime(2026, 1, 5) + timedelta(days=7 * i))
                for i in range(3)]
        self.assertEqual(nomi, ["S02_2026-01-05", "S03_2026-01-12", "S04_2026-01-19"])
        self.assertEqual(sorted(nomi), nomi)

    def test_settimana_a_cavallo_di_due_anni(self):
        """Lunedi' 29-12-2025 e' la settimana ISO 1: va sotto l'anno del lunedi'
        (2025), lo stesso del riepilogo annuale che lo contiene."""
        with tempfile.TemporaryDirectory() as tmp:
            cartella = archivio.cartella_settimana(tmp, datetime(2025, 12, 29))
            self.assertEqual(cartella.relative_to(tmp), Path("2025/S01_2025-12-29"))

    def test_nomi_dei_file(self):
        self.assertEqual(archivio.nome_settimanale(LUNEDI), "Conta_Pollinica_12-10-2026.xlsx")
        self.assertEqual(archivio.nome_annuale(2026), "Riepilogo_Annuale_2026.xlsx")
        self.assertEqual(archivio.nome_bollettino("ITA", LUNEDI), "Bollettino_ITA_12-10-2026.docx")
        self.assertEqual(archivio.nome_bollettino("ENG", LUNEDI), "Bollettino_ENG_12-10-2026.docx")
        self.assertEqual(archivio.nome_journal(LUNEDI), "~sessione_12-10-2026.jsonl")


class TestPercorsi(unittest.TestCase):
    def test_i_cinque_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            atteso = {
                archivio.percorso_settimanale(r, LUNEDI):
                    "2026/S42_2026-10-12/Conta_Pollinica_12-10-2026.xlsx",
                archivio.percorso_annuale(r, 2026):
                    "2026/Riepilogo_Annuale_2026.xlsx",
                archivio.percorso_bollettino(r, "ITA", LUNEDI):
                    "2026/S42_2026-10-12/Bollettino_ITA_12-10-2026.docx",
                archivio.percorso_journal(r, LUNEDI):
                    "in_corso/~sessione_12-10-2026.jsonl",
                archivio.percorso_elenco_codici(r):
                    "Cheatsheet_Codici_Specie.docx",
            }
            for percorso, relativo in atteso.items():
                self.assertEqual(percorso.relative_to(r), Path(relativo))

    def test_prepara_crea_le_cartelle(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertTrue(archivio.prepara_settimana(tmp, LUNEDI).is_dir())
            self.assertTrue(archivio.prepara_anno(tmp, 2026).is_dir())
            self.assertTrue(archivio.prepara_in_corso(tmp).is_dir())
            # idempotente
            self.assertTrue(archivio.prepara_settimana(tmp, LUNEDI).is_dir())

    def test_legge_lunedi_da_nome(self):
        self.assertEqual(archivio.legge_lunedi_da_nome("S42_2026-10-12"), LUNEDI)
        self.assertEqual(archivio.legge_lunedi_da_nome("Conta_Pollinica_12-10-2026.xlsx"), LUNEDI)
        self.assertEqual(archivio.legge_lunedi_da_nome("Bollettino_ENG_12-10-2026.docx"), LUNEDI)
        self.assertEqual(archivio.legge_lunedi_da_nome("~sessione_12-10-2026.jsonl"), LUNEDI)
        # un giorno qualsiasi della settimana viene riportato al suo lunedi'
        self.assertEqual(archivio.legge_lunedi_da_nome("Conta_Pollinica_15-10-2026.xlsx"), LUNEDI)
        # nomi senza data, o con una data impossibile
        self.assertIsNone(archivio.legge_lunedi_da_nome("rifatto_ContaPoll.xlsx"))
        self.assertIsNone(archivio.legge_lunedi_da_nome("Conta_Pollinica_31-02-2026.xlsx"))


class TestElenca(unittest.TestCase):
    def test_cartella_vuota(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(archivio.elenca(tmp), [])

    def test_struttura_nuova(self):
        with tempfile.TemporaryDirectory() as tmp:
            _crea(archivio.percorso_settimanale(tmp, LUNEDI))
            _crea(archivio.percorso_bollettino(tmp, "ITA", LUNEDI))
            _crea(archivio.percorso_bollettino(tmp, "ENG", LUNEDI))
            voci = archivio.elenca(tmp)
            self.assertEqual(len(voci), 1)
            voce = voci[0]
            self.assertEqual(voce.lunedi, LUNEDI)
            self.assertEqual(voce.stato, "salvato")
            self.assertEqual(voce.xlsx.name, archivio.nome_settimanale(LUNEDI))
            self.assertEqual(len(voce.bollettini), 2)
            self.assertFalse(voce.fuori_struttura)

    def test_journal_e_xlsx_della_stessa_settimana_sono_una_voce_sola(self):
        with tempfile.TemporaryDirectory() as tmp:
            _crea(archivio.percorso_settimanale(tmp, LUNEDI))
            _crea(archivio.percorso_journal(tmp, LUNEDI))
            voci = archivio.elenca(tmp)
            self.assertEqual(len(voci), 1)
            self.assertEqual(voci[0].stato, "non_salvato")
            self.assertIsNotNone(voci[0].xlsx)
            self.assertIsNotNone(voci[0].journal)

    def test_journal_nella_vecchia_posizione_resta_visibile(self):
        with tempfile.TemporaryDirectory() as tmp:
            _crea(Path(tmp) / archivio.nome_journal(LUNEDI))
            voci = archivio.elenca(tmp)
            self.assertEqual(len(voci), 1)
            self.assertEqual(voci[0].stato, "non_salvato")

    def test_cartella_piatta_marcata_da_riordinare(self):
        with tempfile.TemporaryDirectory() as tmp:
            _crea(Path(tmp) / archivio.nome_settimanale(LUNEDI))
            voci = archivio.elenca(tmp)
            self.assertEqual(len(voci), 1)
            self.assertTrue(voci[0].fuori_struttura)
            self.assertEqual(voci[0].lunedi, LUNEDI)

    def test_file_messi_da_parte_restano_nell_elenco(self):
        """Quello che il riordino mette in 'da_controllare' non deve sparire
        dall'elenco: se il nome dice la settimana finisce nella sua voce,
        altrimenti compare col proprio nome."""
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            _crea(archivio.percorso_bollettino(r, "ITA", LUNEDI))
            _crea(r / archivio.NOME_DA_CONTROLLARE / "test_Pollencounter_12-10-2026.xlsx")
            _crea(r / archivio.NOME_DA_CONTROLLARE / "rifatto_ContaPoll_160226.xlsx")
            voci = archivio.elenca(r)
            self.assertEqual(len(voci), 2)
            per_settimana = {v.lunedi: v for v in voci}
            self.assertEqual(per_settimana[LUNEDI].xlsx.name, "test_Pollencounter_12-10-2026.xlsx")
            self.assertTrue(per_settimana[LUNEDI].fuori_struttura)
            self.assertEqual(per_settimana[None].etichetta, "rifatto_ContaPoll_160226.xlsx")

    def test_etichetta(self):
        """Il periodo si mostra con l'iniziale maiuscola, il nome di un file
        tale e quale (capitalize() storpiava 'test_Pollencounter...')."""
        voce = archivio.VoceArchivio(lunedi=LUNEDI)
        self.assertTrue(voce.etichetta.startswith("Da luned"))
        strano = archivio.VoceArchivio(xlsx=Path("test_Pollencounter_12-10-2026.xlsx"))
        self.assertEqual(strano.etichetta, "test_Pollencounter_12-10-2026.xlsx")

    def test_file_fuori_schema_passato_da_chi_chiama(self):
        with tempfile.TemporaryDirectory() as tmp:
            strano = _crea(Path(tmp) / "rifatto_ContaPoll_160226.xlsx")
            voci = archivio.elenca(tmp, legacy_xlsx=[strano])
            self.assertEqual(len(voci), 1)
            self.assertIsNone(voci[0].lunedi)
            self.assertTrue(voci[0].fuori_struttura)
            # senza una settimana, nell'elenco si mostra il nome del file
            self.assertEqual(voci[0].etichetta, "rifatto_ContaPoll_160226.xlsx")

    def test_non_apre_i_file(self):
        """L'elenco si costruisce dai nomi: un .xlsx illeggibile compare comunque
        con la sua settimana (prima ogni file veniva aperto con openpyxl)."""
        with tempfile.TemporaryDirectory() as tmp:
            percorso = archivio.percorso_settimanale(tmp, LUNEDI)
            percorso.parent.mkdir(parents=True, exist_ok=True)
            percorso.write_text("non sono un file Excel", encoding="utf-8")
            voci = archivio.elenca(tmp)
            self.assertEqual(len(voci), 1)
            self.assertEqual(voci[0].lunedi, LUNEDI)

    def test_ordine_non_salvati_poi_settimane_recenti(self):
        with tempfile.TemporaryDirectory() as tmp:
            vecchio = datetime(2026, 9, 7)
            _crea(archivio.percorso_settimanale(tmp, LUNEDI))
            _crea(archivio.percorso_settimanale(tmp, ALTRO_LUNEDI))
            _crea(archivio.percorso_settimanale(tmp, vecchio))
            _crea(archivio.percorso_journal(tmp, vecchio))
            voci = archivio.elenca(tmp)
            self.assertEqual([v.lunedi for v in voci], [vecchio, LUNEDI, ALTRO_LUNEDI])

    def test_cartella_di_settimana_vuota_non_e_una_voce(self):
        with tempfile.TemporaryDirectory() as tmp:
            archivio.prepara_settimana(tmp, LUNEDI)
            self.assertEqual(archivio.elenca(tmp), [])


class TestRiordino(unittest.TestCase):
    def test_piano_mette_ogni_file_al_suo_posto(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            _crea(r / archivio.nome_settimanale(LUNEDI))
            _crea(r / archivio.nome_bollettino("ITA", LUNEDI))
            _crea(r / archivio.nome_annuale(2026))
            _crea(r / archivio.nome_journal(LUNEDI))
            piano = dict(archivio.piano_riordino(r))
            atteso = {
                archivio.nome_settimanale(LUNEDI): "2026/S42_2026-10-12",
                archivio.nome_bollettino("ITA", LUNEDI): "2026/S42_2026-10-12",
                archivio.nome_annuale(2026): "2026",
                archivio.nome_journal(LUNEDI): "in_corso",
            }
            for origine, destinazione in piano.items():
                self.assertEqual(destinazione.parent.relative_to(r),
                                 Path(atteso[origine.name]))
                self.assertEqual(destinazione.name, origine.name)

    def test_piano_non_tocca_cio_che_non_e_suo(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            _crea(r / archivio.NOME_ELENCO_CODICI)
            _crea(r / "concentrazioni_polliniche.xlsx")
            _crea(r / "pollencounter.cfg")
            _crea(r / "appunti.txt")
            _crea(r / archivio.NOME_TEMPLATE)
            _crea(r / f"{archivio.PREFISSO_RESOCONTO}2026-10-16_1200.txt")
            # le soglie personalizzate vengono lette dalla radice: spostarle le
            # renderebbe invisibili
            self.assertEqual(archivio.piano_riordino(r), [])

    def test_file_fuori_schema_in_da_controllare(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            _crea(r / "rifatto_ContaPoll_160226.xlsx")
            (origine, destinazione), = archivio.piano_riordino(r)
            self.assertEqual(destinazione.relative_to(r),
                             Path(archivio.NOME_DA_CONTROLLARE) / origine.name)

    def test_esegui_sposta_e_scrive_il_resoconto(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            origine = _crea(r / archivio.nome_settimanale(LUNEDI))
            origine.write_text("dati", encoding="utf-8")
            spostati, falliti = archivio.esegui_riordino(archivio.piano_riordino(r), r)
            self.assertEqual(falliti, [])
            self.assertEqual(len(spostati), 1)
            self.assertFalse(origine.exists())
            arrivo = archivio.percorso_settimanale(r, LUNEDI)
            self.assertEqual(arrivo.read_text(encoding="utf-8"), "dati")
            resoconti = list(r.glob(f"{archivio.PREFISSO_RESOCONTO}*.txt"))
            self.assertEqual(len(resoconti), 1)
            self.assertIn(archivio.nome_settimanale(LUNEDI),
                          resoconti[0].read_text(encoding="utf-8"))

    def test_esegui_non_sovrascrive_mai(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            origine = _crea(r / archivio.nome_settimanale(LUNEDI))
            origine.write_text("vecchio", encoding="utf-8")
            arrivo = _crea(archivio.percorso_settimanale(r, LUNEDI))
            arrivo.write_text("nuovo", encoding="utf-8")

            spostati, falliti = archivio.esegui_riordino(archivio.piano_riordino(r), r)
            self.assertEqual(spostati, [])
            self.assertEqual(len(falliti), 1)
            self.assertIn("esiste gia'", falliti[0][1])
            # nessuno dei due file e' stato toccato
            self.assertEqual(origine.read_text(encoding="utf-8"), "vecchio")
            self.assertEqual(arrivo.read_text(encoding="utf-8"), "nuovo")

    def test_serve_riordino(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertFalse(archivio.serve_riordino(tmp))
            _crea(Path(tmp) / archivio.nome_settimanale(LUNEDI))
            self.assertTrue(archivio.serve_riordino(tmp))


class TestArchiviaOriginale(unittest.TestCase):
    def test_sposta_il_file_di_partenza_nella_sua_settimana(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            origine = _crea(r / "rifatto_ContaPoll_160226.xlsx")
            destinazione = archivio.archivia_originale(r, LUNEDI, origine)
            self.assertEqual(destinazione.parent, archivio.cartella_settimana(r, LUNEDI))
            self.assertFalse(origine.exists())

    def test_nome_uguale_al_file_appena_salvato_non_lo_sovrascrive(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            origine = _crea(r / archivio.nome_settimanale(LUNEDI))
            origine.write_text("vecchio", encoding="utf-8")
            salvato = _crea(archivio.percorso_settimanale(r, LUNEDI))
            salvato.write_text("appena salvato", encoding="utf-8")

            destinazione = archivio.archivia_originale(r, LUNEDI, origine)
            self.assertEqual(destinazione.name, f"originale_{archivio.nome_settimanale(LUNEDI)}")
            self.assertEqual(salvato.read_text(encoding="utf-8"), "appena salvato")
            self.assertEqual(destinazione.read_text(encoding="utf-8"), "vecchio")

    def test_niente_da_fare(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = Path(tmp)
            self.assertIsNone(archivio.archivia_originale(r, LUNEDI, None))
            self.assertIsNone(archivio.archivia_originale(r, LUNEDI, r / "mai_esistito.xlsx"))
            # un file gia' nell'archivio non si sposta
            dentro = _crea(archivio.percorso_settimanale(r, LUNEDI))
            self.assertIsNone(archivio.archivia_originale(r, LUNEDI, dentro))
            self.assertTrue(dentro.exists())


if __name__ == "__main__":
    unittest.main()
