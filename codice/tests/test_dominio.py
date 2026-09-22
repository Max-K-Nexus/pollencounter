import unittest
from datetime import datetime

import dominio as d


class TestCodici(unittest.TestCase):
    def test_normalizza_codice_singola_cifra(self):
        self.assertEqual(d.normalizza_codice("5"), "05")

    def test_normalizza_codice_gia_normalizzato(self):
        self.assertEqual(d.normalizza_codice("48"), "48")

    def test_codice_valido(self):
        self.assertTrue(d.codice_valido("01"))
        self.assertTrue(d.codice_valido("59"))
        self.assertFalse(d.codice_valido("00"))
        self.assertFalse(d.codice_valido("60"))

    def test_tutti_i_codici_soglie_mapping_sono_validi(self):
        for codice in d.SOGLIE_MAPPING:
            self.assertIn(codice, d.CODICI_SPECIE)

    def test_tutti_i_codici_bollettino_sono_validi(self):
        for _, _, codici, famiglia in d.BOLLETTINO_RIGHE:
            for codice in codici:
                self.assertIn(codice, d.CODICI_SPECIE)
            self.assertIn(famiglia, d.SOGLIE_FALLBACK)


class TestDate(unittest.TestCase):
    def test_formato_numerico_trattino(self):
        self.assertEqual(d.parse_data_flessibile("9-2-2026"), datetime(2026, 2, 9))

    def test_formato_numerico_slash_anno_corto(self):
        self.assertEqual(d.parse_data_flessibile("9/2/26"), datetime(2026, 2, 9))

    def test_formato_testuale_mese_esteso(self):
        self.assertEqual(d.parse_data_flessibile("9 febbraio 2026"), datetime(2026, 2, 9))

    def test_formato_testuale_mese_abbreviato(self):
        self.assertEqual(d.parse_data_flessibile("9 feb 2026"), datetime(2026, 2, 9))

    def test_testo_non_riconosciuto(self):
        self.assertIsNone(d.parse_data_flessibile("qualcosa"))

    def test_data_invalida(self):
        self.assertIsNone(d.parse_data_flessibile("32-13-2026"))

    def test_lunedi_di(self):
        # 9 feb 2026 e' un lunedi'
        self.assertEqual(d.lunedi_di(datetime(2026, 2, 9)), datetime(2026, 2, 9))
        # 12 feb 2026 e' un giovedi' della stessa settimana
        self.assertEqual(d.lunedi_di(datetime(2026, 2, 12)), datetime(2026, 2, 9))


class TestConcentrazione(unittest.TestCase):
    def test_conta_zero_ritorna_none(self):
        self.assertIsNone(d.concentrazione(0, 0.4))

    def test_calcolo_base(self):
        self.assertEqual(d.concentrazione(10, 0.4), 4.0)

    def test_arrotondamento(self):
        self.assertEqual(d.concentrazione(3, 0.4), 1.2)


class TestSoglie(unittest.TestCase):
    def test_parse_soglia_max_range(self):
        self.assertEqual(d.parse_soglia_max("0 - 0,5"), 0.5)

    def test_parse_soglia_max_minore(self):
        self.assertAlmostEqual(d.parse_soglia_max("< 100"), 99.9)

    def test_parse_soglia_max_maggiore(self):
        self.assertEqual(d.parse_soglia_max("> 50"), float("inf"))

    def test_parse_soglia_max_numero_singolo(self):
        self.assertEqual(d.parse_soglia_max("25"), 25.0)

    def test_parse_soglia_max_vuoto(self):
        self.assertEqual(d.parse_soglia_max(""), 0.0)
        self.assertEqual(d.parse_soglia_max(None), 0.0)

    def test_parse_soglie_salta_righe_titolo(self):
        righe = [
            ("Spore Fungine", None, None, None),
            ("Alternaria", "0 - 1,9", "2 - 19", "20 - 100"),
        ]
        soglie = d.parse_soglie(righe)
        self.assertNotIn("Spore Fungine", soglie)
        self.assertEqual(soglie["Alternaria"], (1.9, 19.0, 100.0))

    def test_livello(self):
        soglia = (0.9, 19.9, 39.9)
        self.assertEqual(d.livello(0.0, soglia), "assente")
        self.assertEqual(d.livello(0.9, soglia), "assente")
        self.assertEqual(d.livello(1.0, soglia), "bassa")
        self.assertEqual(d.livello(20.0, soglia), "media")
        self.assertEqual(d.livello(40.0, soglia), "alta")
        self.assertEqual(d.livello(None, soglia), "assente")


class TestRigheBollettino(unittest.TestCase):
    def test_specie_senza_dati_esclusa(self):
        righe = d.righe_bollettino({}, 0.4, {})
        self.assertEqual(righe, [])

    def test_specie_semplice(self):
        conteggi = {"01": [1, 0, 0, 0, 0, 0, 0]}
        righe = d.righe_bollettino(conteggi, 0.4, {})
        self.assertEqual(len(righe), 1)
        r = righe[0]
        self.assertEqual(r.nome_ita, "ACERACEAE")
        self.assertEqual(r.concentrazioni[0], 0.4)
        self.assertEqual(r.media, round(0.4 / 7, 1))
        self.assertEqual(len(r.livelli_giorno), 7)

    def test_specie_aggregata_somma_i_codici(self):
        # BETULACEAE = 03 + 04 + 05
        conteggi = {
            "03": [1, 0, 0, 0, 0, 0, 0],
            "04": [2, 0, 0, 0, 0, 0, 0],
            "05": [0, 0, 0, 0, 0, 0, 0],
        }
        righe = d.righe_bollettino(conteggi, 1.0, {})
        nomi = {r.nome_ita: r for r in righe}
        self.assertIn("BETULACEAE", nomi)
        self.assertEqual(nomi["BETULACEAE"].concentrazioni[0], 3.0)
        # anche Alnus (04) appare come riga individuale
        self.assertIn("Alnus", nomi)
        self.assertEqual(nomi["Alnus"].concentrazioni[0], 2.0)

    def test_soglie_esterne_hanno_precedenza_sul_fallback(self):
        conteggi = {"48": [300, 0, 0, 0, 0, 0, 0]}  # Alternaria
        soglie = {"Alternaria": (1.9, 19.0, 500.0)}  # soglia "media" molto alta
        righe = d.righe_bollettino(conteggi, 1.0, soglie)
        # media = 300/7 = 42.9, sotto i 500 di soglia media esterna -> "media"
        self.assertEqual(righe[0].livello_media, "media")

    def test_livello_per_giorno_indipendente_dalla_media(self):
        # lunedi molto alto, resto della settimana a zero: la media resta
        # bassa ma il lunedi' deve comunque risultare "alta" (difetto #3)
        conteggi = {"48": [500, 0, 0, 0, 0, 0, 0]}
        soglie = {"Alternaria": (1.9, 19.0, 100.0)}
        righe = d.righe_bollettino(conteggi, 1.0, soglie)
        r = righe[0]
        self.assertEqual(r.livelli_giorno[0], "alta")
        self.assertNotEqual(r.livelli_giorno[0], r.livello_media)


class TestInterpretaComando(unittest.TestCase):
    def test_codice_semplice(self):
        r = d.interpreta_comando("05")
        self.assertIsInstance(r, d.Inserisci)
        self.assertEqual(r.codice, "05")
        self.assertEqual(r.quantita, 1)

    def test_codice_singola_cifra_normalizzato(self):
        r = d.interpreta_comando("5")
        self.assertEqual(r.codice, "05")

    def test_moltiplicatore(self):
        r = d.interpreta_comando("48x4")
        self.assertIsInstance(r, d.Inserisci)
        self.assertEqual(r.codice, "48")
        self.assertEqual(r.quantita, 4)

    def test_moltiplicatore_maiuscolo_e_asterisco(self):
        self.assertEqual(d.interpreta_comando("48X4").quantita, 4)
        self.assertEqual(d.interpreta_comando("48*4").quantita, 4)

    def test_moltiplicatore_fuori_range(self):
        r = d.interpreta_comando("48x101")
        self.assertIsInstance(r, d.ComandoNonValido)

    def test_ripeti(self):
        self.assertIsInstance(d.interpreta_comando("."), d.Ripeti)

    def test_azione_singola_lettera(self):
        r = d.interpreta_comando("u")
        self.assertIsInstance(r, d.Azione)
        self.assertEqual(r.lettera, "u")

    def test_codice_non_riconosciuto(self):
        r = d.interpreta_comando("99")
        self.assertIsInstance(r, d.ComandoNonValido)

    def test_stringa_vuota(self):
        r = d.interpreta_comando("")
        self.assertIsInstance(r, d.ComandoNonValido)


if __name__ == "__main__":
    unittest.main()
