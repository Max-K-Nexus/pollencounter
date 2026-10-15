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


class TestDescrizioniGui(unittest.TestCase):
    def test_nome_giorno_accentato(self):
        self.assertEqual(d.nome_giorno_accentato(datetime(2026, 10, 15)), "giovedì")
        self.assertEqual(d.nome_giorno_accentato(datetime(2026, 10, 12)), "lunedì")
        self.assertEqual(d.nome_giorno_accentato(datetime(2026, 10, 18)), "domenica")

    def test_descrivi_settimana(self):
        self.assertEqual(d.descrivi_settimana(datetime(2026, 10, 12)),
                         "da lunedì 12/10 a domenica 18/10/2026")

    def test_descrivi_settimana_a_cavallo_d_anno(self):
        self.assertEqual(d.descrivi_settimana(datetime(2026, 12, 28)),
                         "da lunedì 28/12 a domenica 03/01/2027")

    def test_descrivi_quando(self):
        adesso = datetime(2026, 10, 15, 18, 0)
        self.assertEqual(d.descrivi_quando(datetime(2026, 10, 15, 14, 32), adesso), "oggi 14:32")
        self.assertEqual(d.descrivi_quando(datetime(2026, 10, 14, 9, 10), adesso), "ieri 09:10")
        self.assertEqual(d.descrivi_quando(datetime(2026, 10, 12, 9, 10), adesso), "12/10/2026")

    def test_descrivi_quando_dopo_mezzanotte(self):
        # 23:50 di ieri visto alle 00:10 di oggi: "ieri", non "oggi"
        self.assertEqual(d.descrivi_quando(datetime(2026, 10, 14, 23, 50),
                                           datetime(2026, 10, 15, 0, 10)), "ieri 23:50")


class TestConcentrazione(unittest.TestCase):
    def test_conta_zero_ritorna_none(self):
        self.assertIsNone(d.concentrazione(0, 0.4))

    def test_calcolo_base(self):
        self.assertEqual(d.concentrazione(10, 0.4), 4.0)

    def test_arrotondamento(self):
        self.assertEqual(d.concentrazione(3, 0.4), 1.2)


class TestInterpretaFattore(unittest.TestCase):
    def test_virgola_e_punto(self):
        self.assertEqual(d.interpreta_fattore("0,5"), 0.5)
        self.assertEqual(d.interpreta_fattore("0.5"), 0.5)
        self.assertEqual(d.interpreta_fattore(" 1 "), 1.0)

    def test_valori_non_validi(self):
        for testo in ("0", "-1", "abc", "", "1000", "nan", "inf"):
            with self.subTest(testo=testo):
                with self.assertRaises(ValueError):
                    d.interpreta_fattore(testo)

    def test_messaggio_in_italiano_comprensibile(self):
        with self.assertRaises(ValueError) as ctx:
            d.interpreta_fattore("abc")
        self.assertIn("non e' un numero", str(ctx.exception))

    def test_lettera_f_e_un_comando(self):
        self.assertEqual(d.interpreta_comando("f"), d.Azione("f"))

    def test_il_fattore_non_e_comandabile_a_voce(self):
        azioni = {c.lettera for c in d.COMANDI_VOCALI.values() if isinstance(c, d.Azione)}
        self.assertNotIn("f", azioni)


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


class TestSogliaFamiglia(unittest.TestCase):
    def test_dal_file_poi_fallback_poi_default(self):
        self.assertEqual(d.soglia_famiglia({"Aceracee": (1, 2, 3)}, "Aceracee"), (1, 2, 3))
        self.assertEqual(d.soglia_famiglia({}, "Aceracee"), d.SOGLIE_FALLBACK["Aceracee"])
        self.assertEqual(d.soglia_famiglia({}, "Sconosciuta"), (0.9, 19.9, 39.9))


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


class TestVocale(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.vocab, cls.avvisi = d.costruisci_vocabolario_vocale()

    def v(self, frase):
        """Interpreta una frase pronunciata con la parola di attivazione."""
        return d.interpreta_vocale("conta " + frase if frase else frase, self.vocab)

    def test_numeri_in_parole(self):
        self.assertEqual(d.numero_in_parole(1), "uno")
        self.assertEqual(d.numero_in_parole(21), "ventuno")
        self.assertEqual(d.numero_in_parole(28), "ventotto")
        self.assertEqual(d.numero_in_parole(33), "trentatre")
        self.assertEqual(d.numero_in_parole(50), "cinquanta")
        self.assertEqual(d.numero_in_parole(59), "cinquantanove")

    def test_numero_in_parole_fuori_range(self):
        with self.assertRaises(ValueError):
            d.numero_in_parole(0)
        with self.assertRaises(ValueError):
            d.numero_in_parole(60)

    def test_vocabolario_di_default_senza_conflitti(self):
        self.assertEqual(self.avvisi, [])

    def test_ogni_codice_ha_numero_e_nome_parlati(self):
        for codice, nome in d.CODICI_SPECIE.items():
            self.assertEqual(self.vocab[d.numero_in_parole(int(codice))], codice)
            self.assertEqual(self.vocab[d.normalizza_parlato(nome)], codice)

    def test_codice_a_voce(self):
        r = self.v("ventiquattro")
        self.assertEqual((r.codice, r.quantita), ("24", 1))

    def test_nome_a_voce_con_sinonimo(self):
        self.assertEqual(self.v("Graminacee").codice, "24")
        self.assertEqual(self.v("gramineae").codice, "24")

    def test_nome_con_parentesi_e_barra(self):
        self.assertEqual(self.v("corylaceae").codice, "12")
        self.assertEqual(self.v("carpinus ostrya").codice, "13")

    def test_quantita_a_voce(self):
        r = self.v("graminacee per tre")
        self.assertEqual((r.codice, r.quantita), ("24", 3))
        r = self.v("uno per venti")
        self.assertEqual((r.codice, r.quantita), ("01", 20))

    def test_quantita_fuori_range_o_non_capita(self):
        for frase in ("graminacee per ventuno", "graminacee per pippo"):
            self.assertIsInstance(self.v(frase), d.ComandoNonValido)

    def test_per_uno_vale_una_occorrenza(self):
        r = self.v("graminacee per uno")
        self.assertEqual((r.codice, r.quantita), ("24", 1))
        self.assertIn("conta acero per uno", d.costruisci_grammatica(self.vocab))

    def test_sequenza_di_comandi_in_una_frase(self):
        seq = d.interpreta_vocale_sequenza(
            "conta acero conta ontano conta graminacee per tre", self.vocab)
        self.assertEqual([s for s, _ in seq],
                         ["conta acero", "conta ontano", "conta graminacee per tre"])
        self.assertEqual([c for _, c in seq],
                         [d.Inserisci("01", 1), d.Inserisci("04", 1), d.Inserisci("24", 3)])

    def test_sequenza_con_annulla_in_ordine(self):
        seq = d.interpreta_vocale_sequenza("conta acero conta annulla conta ontano", self.vocab)
        self.assertEqual([c for _, c in seq],
                         [d.Inserisci("01", 1), d.Azione("u"), d.Inserisci("04", 1)])

    def test_sequenza_di_un_solo_comando_come_prima(self):
        seq = d.interpreta_vocale_sequenza("conta acero per uno", self.vocab)
        self.assertEqual(seq, [("conta acero per uno", d.Inserisci("01", 1))])

    def test_sequenza_comando_non_valido_non_blocca_gli_altri(self):
        seq = d.interpreta_vocale_sequenza(
            "conta acero conta [unk] conta pippo conta ontano", self.vocab)
        comandi = [c for _, c in seq]
        self.assertEqual(comandi[0], d.Inserisci("01", 1))
        self.assertIsInstance(comandi[1], d.ComandoNonValido)
        self.assertIsInstance(comandi[2], d.ComandoNonValido)
        self.assertEqual(comandi[3], d.Inserisci("04", 1))

    def test_dividi_scarta_cio_che_precede_la_prima_attivazione(self):
        self.assertEqual(d.dividi_comandi_vocali("acero tre conta ontano"), ["conta ontano"])
        self.assertEqual(d.dividi_comandi_vocali("acero tre"), [])
        self.assertEqual(d.dividi_comandi_vocali(""), [])

    def test_dividi_ignora_attivazioni_a_vuoto(self):
        self.assertEqual(d.dividi_comandi_vocali("conta conta acero conta"), ["conta acero"])

    def test_dividi_conserva_unk_e_normalizza(self):
        self.assertEqual(d.dividi_comandi_vocali("Conta [unk] CONTA Acero"),
                         ["conta [unk]", "conta acero"])

    def test_dividi_senza_parola_di_attivazione_e_un_solo_comando(self):
        self.assertEqual(d.dividi_comandi_vocali("acero per tre", ""), ["acero per tre"])
        seq = d.interpreta_vocale_sequenza("acero per tre", self.vocab, attivazione="")
        self.assertEqual([c for _, c in seq], [d.Inserisci("01", 3)])

    def test_dividi_attivazione_personalizzata(self):
        self.assertEqual(d.dividi_comandi_vocali("registra acero registra ontano", "registra"),
                         ["registra acero", "registra ontano"])

    def test_comandi_vocali_reversibili(self):
        self.assertEqual(self.v("annulla"), d.Azione("u"))
        self.assertIsInstance(self.v("ripeti"), d.Ripeti)
        self.assertIsInstance(self.v("ancora"), d.Ripeti)
        self.assertIsInstance(self.v("totale"), d.Totale)

    def test_nessun_comando_distruttivo_raggiungibile_a_voce(self):
        azioni = {c.lettera for c in d.COMANDI_VOCALI.values() if isinstance(c, d.Azione)}
        self.assertEqual(azioni, {"u"})
        for frase in ("salva", "chiudi", "esci", "chiudi giornata", "s", "q", "d"):
            self.assertNotIsInstance(self.v(frase), d.Azione)

    def test_fuori_vocabolario_scartato_in_silenzio(self):
        for frase in ("[unk]", "ventiquattro [unk]", "buongiorno a tutti", ""):
            r = self.v(frase)
            self.assertIsInstance(r, d.ComandoNonValido)
            self.assertEqual(r.messaggio, "")

    def test_grammatica_coerente_con_interprete(self):
        for frase in d.costruisci_grammatica(self.vocab):
            self.assertNotIsInstance(d.interpreta_vocale(frase, self.vocab), d.ComandoNonValido, frase)
        for frase in d.costruisci_grammatica(self.vocab, attivazione=""):
            self.assertNotIsInstance(d.interpreta_vocale(frase, self.vocab, attivazione=""),
                                     d.ComandoNonValido, frase)

    def test_sinonimi_personali(self):
        vocab, avvisi = d.costruisci_vocabolario_vocale({"24": ["erba"]})
        self.assertEqual(d.interpreta_vocale("conta erba", vocab).codice, "24")
        self.assertEqual(avvisi, [])

    def test_sinonimo_personale_in_conflitto_segnalato(self):
        vocab, avvisi = d.costruisci_vocabolario_vocale({"05": ["graminacee"]})
        self.assertEqual(d.interpreta_vocale("conta graminacee", vocab).codice, "24")
        self.assertEqual(len(avvisi), 1)

    def test_sinonimo_personale_su_comando_ignorato(self):
        vocab, avvisi = d.costruisci_vocabolario_vocale({"24": ["annulla"]})
        self.assertNotIn("annulla", vocab)
        self.assertEqual(len(avvisi), 1)

    def test_sinonimi_per_codice_inesistente_segnalati(self):
        _, avvisi = d.costruisci_vocabolario_vocale({"99": ["pippo"]})
        self.assertEqual(len(avvisi), 1)


    def test_senza_parola_di_attivazione_scartato(self):
        # "due", "tre", "sei" sono parole comuni: senza "conta" non registrano nulla
        for frase in ("due", "tre", "sei", "ventiquattro", "graminacee per tre", "annulla", "totale"):
            r = d.interpreta_vocale(frase, self.vocab)
            self.assertIsInstance(r, d.ComandoNonValido, frase)
            self.assertEqual(r.messaggio, "")

    def test_parola_di_attivazione_sola_scartata(self):
        self.assertIsInstance(d.interpreta_vocale("conta", self.vocab), d.ComandoNonValido)

    def test_prefisso_disattivato(self):
        r = d.interpreta_vocale("ventiquattro", self.vocab, attivazione="")
        self.assertEqual((r.codice, r.quantita), ("24", 1))
        self.assertIsInstance(d.interpreta_vocale("annulla", self.vocab, attivazione=""), d.Azione)

    def test_parola_di_attivazione_personalizzata(self):
        r = d.interpreta_vocale("registra graminacee", self.vocab, attivazione="registra")
        self.assertEqual(r.codice, "24")
        self.assertIsInstance(d.interpreta_vocale("conta graminacee", self.vocab, attivazione="registra"),
                              d.ComandoNonValido)

    def test_grammatica_tutta_con_prefisso(self):
        grammatica = d.costruisci_grammatica(self.vocab)
        self.assertTrue(all(f.startswith("conta ") for f in grammatica))
        self.assertIn("conta ventiquattro per tre", grammatica)

    def test_sinonimo_uguale_alla_parola_di_attivazione_ignorato(self):
        vocab, avvisi = d.costruisci_vocabolario_vocale({"24": ["conta"]})
        self.assertNotIn("conta", vocab)
        self.assertEqual(len(avvisi), 1)


class TestConfermeNonSembranoComandi(unittest.TestCase):
    """Con il microfono sempre aperto mentre il PC parla (anti-eco spento), la
    voce del PC non deve poter essere scambiata per un comando: nessuna
    conferma parlata puo' iniziare con la parola di attivazione."""

    def test_nomi_delle_specie(self):
        for codice, nome in d.CODICI_SPECIE.items():
            parlato = d.normalizza_parlato(nome)
            self.assertFalse(parlato.startswith(d.PAROLA_ATTIVAZIONE_DEFAULT),
                             f"{codice}: {parlato}")
            self.assertNotIn(" " + d.PAROLA_ATTIVAZIONE_DEFAULT + " ", " " + parlato + " ")

    def test_frasi_fisse(self):
        for frase in ("annullato", "totale 12", "non capito", "niente da ripetere", "voce attiva"):
            self.assertFalse(frase.startswith(d.PAROLA_ATTIVAZIONE_DEFAULT))

    def test_una_conferma_non_viene_interpretata_come_comando(self):
        vocabolario, _ = d.costruisci_vocabolario_vocale({}, d.PAROLA_ATTIVAZIONE_DEFAULT)
        for codice in d.CODICI_SPECIE:
            conferma = d.normalizza_parlato(d.CODICI_SPECIE[codice])
            cmd = d.interpreta_vocale(conferma, vocabolario, d.PAROLA_ATTIVAZIONE_DEFAULT)
            self.assertIsInstance(cmd, d.ComandoNonValido, conferma)


class TestPronunceVocali(unittest.TestCase):
    def setUp(self):
        self.voc, self.avvisi = d.costruisci_vocabolario_vocale()

    def test_nessun_avviso_di_collisione(self):
        self.assertEqual(self.avvisi, [])

    def test_ogni_trascrizione_porta_al_suo_codice(self):
        for nome, (codice, trascrizioni) in d.PRONUNCE_VOCALI.items():
            self.assertTrue(d.codice_valido(codice), nome)
            self.assertTrue(trascrizioni, nome)
            for t in trascrizioni:
                cmd = d.interpreta_vocale(f"conta {t}", self.voc)
                self.assertEqual(cmd, d.Inserisci(codice, 1), f"{nome}: {t}")

    def test_trascrizione_con_quantita(self):
        self.assertEqual(d.interpreta_vocale("conta alterna ria per tre", self.voc),
                         d.Inserisci("48", 3))

    def test_trascrizioni_sono_frasi_di_almeno_cinque_lettere_e_senza_per(self):
        for t in d.TRASCRIZIONI_VOCALI:
            self.assertGreaterEqual(len(t.replace(" ", "")), 5, t)
            self.assertNotIn("per", t.split(), t)
            self.assertEqual(t, d.normalizza_parlato(t))

    def test_trascrizione_ricondotta_al_nome_leggibile(self):
        self.assertEqual(d.TRASCRIZIONI_VOCALI["bet ula cee"], "betulacee")
        self.assertIn("betulacee", d.PRONUNCE_VOCALI)

    def test_trascrizioni_nella_grammatica(self):
        grammatica = d.costruisci_grammatica(self.voc)
        self.assertIn("conta bot rite", grammatica)
        self.assertIn("conta bot rite per due", grammatica)

    def test_nomi_comuni_aggiunti(self):
        for frase, codice in (("cup", "17"), ("olio", "35"), ("carpino", "14"), ("lauro", "27"),
                              ("mora", "29"), ("pinus", "36"), ("poligoni", "40"),
                              ("altri olivi", "32"), ("non identificato", "39")):
            self.assertEqual(d.interpreta_vocale(f"conta {frase}", self.voc),
                             d.Inserisci(codice, 1), frase)

    def test_i_refusi_dell_elenco_non_sono_stati_inseriti(self):
        for refuso in ("ippocastanao", "venticinqu", "sentisette", "trentase"):
            self.assertNotIn(refuso, self.voc)

    def test_trentotto_resta_di_platanaceae(self):
        self.assertEqual(d.interpreta_vocale("conta trentotto", self.voc), d.Inserisci("38", 1))


class TestFamiglieSomma(unittest.TestCase):
    def test_le_sei_famiglie_indicate_dall_utente(self):
        self.assertEqual(d.FAMIGLIE_SOMMA, {
            "03": ("03", "04", "05"),               # Betulaceae = Alnus + Betula
            "08": ("08", "09", "10", "11"),         # Compositae
            "12": ("12", "13", "14", "15", "16"),   # Corylaceae
            "20": ("20", "21", "22", "23"),         # Fagaceae
            "31": ("31", "32", "33", "34", "35"),   # Oleaceae
            "41": ("41", "42", "43"),               # Salicaceae
        })

    def test_somma_dubbi_e_specie(self):
        conteggi = {"03": [1, 0, 0, 0, 0, 0, 0], "04": [2, 0, 0, 0, 0, 0, 3],
                    "05": [0, 4, 0, 0, 0, 0, 0]}
        self.assertEqual(d.totali_famiglia(conteggi)["03"], [3, 4, 0, 0, 0, 0, 3])

    def test_famiglie_a_zero_assenti(self):
        conteggi = {"24": [5] * 7}                  # graminacee: nessuna somma
        self.assertEqual(d.totali_famiglia(conteggi), {})

    def test_codici_mancanti_non_danno_errore(self):
        self.assertEqual(d.totali_famiglia({"43": [0, 0, 2, 0, 0, 0, 0]})["41"],
                         [0, 0, 2, 0, 0, 0, 0])


if __name__ == "__main__":
    unittest.main()
