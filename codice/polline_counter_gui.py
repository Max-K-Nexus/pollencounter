#!/usr/bin/env python3
"""
Conta Pollinica — interfaccia grafica.

Processo unico: la GUI chiama direttamente sessione.py/dominio.py/esportatori.py,
lo stesso dominio usato dalla CLI. Ogni inserimento aggiorna il modello in
memoria e notifica le tab con una chiamata diretta (nessun polling, nessuna
corsa dati).
"""

import os
import re
import subprocess
import sys
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

import dominio
import esportatori
import percorsi
import sessione
import voce

try:
    import openpyxl  # noqa: F401  (verificato qui per il messaggio d'errore in main())
except ImportError:
    openpyxl = None

try:
    import sv_ttk
except ImportError:
    sv_ttk = None

if sys.platform == "win32":
    _MONO_FONT = "Courier New"
elif sys.platform == "darwin":
    _MONO_FONT = "Menlo"
else:
    _MONO_FONT = "Monospace"

# Font dell'interfaccia: proporzionale (quello di sistema, "Segoe UI" su Windows).
# Si risolve in _imposta_tema(), quando esiste la finestra; _MONO_FONT resta
# solo per la casella dei codici.
_UI_FONT = "Segoe UI" if sys.platform == "win32" else "TkDefaultFont"

# Palette (Windows 11, tema chiaro). Un solo posto per i colori che non vengono
# dal tema: registro, avvisi, stato.
_COLORI = {
    "accento": "#0067C0",
    "ok": "#0F7B0F",
    "avviso": "#9D5D00",
    "errore": "#C42B1C",
    "grigio": "#605E5C",
    "riquadro_avviso": "#FFF4CE",
    "sfondo_registro": "#FFFFFF",
    "testo_registro": "#1B1B1B",
    "famiglia": "#EAEFF7",
}

MAX_LINES_LOG = 3000
_MIN_SINISTRA, _MIN_DESTRA = 330, 410     # larghezze minime dei due pannelli (px)

HELP_TEXT = """Codici: 01-59 inserisce la specie corrispondente.
NNxQ inserisce Q occorrenze (es. 48x4 = 4 Alternaria).
.    ripete l'ultimo codice inserito.

Comandi (lettera + Invio nella casella di inserimento):
  r   riepilogo giornata corrente (apre la scheda Giornaliero)
  w   riepilogo settimanale (apre la scheda Settimanale)
  l   ultimi inserimenti
  c   correggi un giorno precedente
  n   aggiungi una nota per la giornata
  f   cambia il fattore di conversione (scheda Bollettino: pulsante "Cambia fattore")
  u   annulla ultimo inserimento
  b   attiva/disattiva il beep sonoro
  s   salva il file (senza uscire)
  d   chiudi la giornata corrente
  q   salva ed esci

Gli stessi comandi sono disponibili anche dai pulsanti sulla sinistra e dai
menu in alto. Scorciatoie: Ctrl+S salva, Ctrl+Z annulla, Ctrl+Q esce dal
programma, Ctrl+1...4 cambia scheda, F1 apre questo aiuto.

Lettura vocale (pulsante "Voce"): ogni frase inizia con "conta".
  conta ventiquattro        inserisce il codice 24
  conta graminacee          inserisce la specie per nome
  conta graminacee per tre  inserisce 3 occorrenze (da 1 a 20)
  conta annulla             annulla l'ultimo inserimento
  conta ripeti              ripete l'ultimo codice
  conta totale              il computer legge il totale del giorno
Piu' comandi di seguito: ogni "conta" chiude il precedente e ne inizia uno
nuovo (es. "conta acero conta ontano conta graminacee per tre" registra tre
inserimenti). Fra un comando e l'altro basta una pausa breve.
Con "Ripeti a voce" attivo il computer ripete il nome di cio' che ha
registrato (es. "conta alternaria" -> "alternaria"): utile per controllare
senza staccare gli occhi dal microscopio. "conta totale" risponde sempre.
Salvare, chiudere la giornata e uscire restano solo da tastiera/pulsante."""

HELP_AVVIO = """Come si comincia

Settimana nuova: controlla la data (le frecce passano alla settimana
precedente o successiva) e premi "Inizia nuova settimana".

Continuare un lavoro: scegli la riga nell'elenco e premi "Riprendi
selezionato" (oppure doppio clic).
  Non salvato  = il programma si e' chiuso prima del salvataggio: i dati sono
                 al sicuro e si possono riprendere da dove eri rimasto.
                 Se non servono piu' (prove, settimane abbandonate) usa
                 "Scegli o rimuovi..." nel riquadro giallo, o "Rimuovi non
                 salvato": la rimozione e' definitiva e chiede conferma.
  File salvato = un file Excel gia' salvato, a cui aggiungere altri dati.

I file stanno nella cartella di lavoro indicata in basso ("Cambia..." per
sceglierne un'altra, "Apri" per vederla in Esplora file).
Per chiudere il programma premi "Esci" (o Ctrl+Q)."""


def _sfondo_tema():
    """Colore di sfondo del tema in uso, per i widget tk che non lo seguono da soli."""
    return ttk.Style().lookup("TFrame", "background") or "#ffffff"


def _imposta_tema(root):
    """Tema Windows 11 (sv_ttk) se installato, altrimenti un ripiego pulito con gli
    stessi nomi di stile ("Accent.TButton"), cosi' il resto del codice non cambia.
    Risolve anche il font dell'interfaccia."""
    global _UI_FONT
    style = ttk.Style(root)
    if sv_ttk is not None:
        sv_ttk.set_theme("light")
    else:
        if sys.platform != "win32" and "clam" in style.theme_names():
            style.theme_use("clam")      # su Windows resta il tema nativo
        style.configure("Accent.TButton", background=_COLORI["accento"], foreground="#FFFFFF",
                        font=(_UI_FONT, 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#1A7FD4"), ("disabled", "#B0B0B0")],
                  foreground=[("disabled", "#F0F0F0")])
    if _UI_FONT == "TkDefaultFont":
        _UI_FONT = tkfont.nametofont("TkDefaultFont").actual("family")
    # Dimensione leggibile per chi lavora al microscopio e guarda lo schermo di sfuggita.
    for nome in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
        try:
            tkfont.nametofont(nome).configure(size=10)
        except tk.TclError:
            pass
    style.configure("Giorno.TButton", padding=(2, 5))
    style.configure("Giorno.Accent.TButton", padding=(2, 5))
    style.configure("TNotebook.Tab", font=(_UI_FONT, 9))
    style.configure("Titolo.TLabel", font=(_UI_FONT, 20, "bold"))
    style.configure("Sezione.TLabel", font=(_UI_FONT, 11, "bold"))
    style.configure("Giorno.TLabel", font=(_UI_FONT, 12, "bold"))
    style.configure("Summary.Treeview", font=(_UI_FONT, 10), rowheight=24)
    style.configure("Summary.Treeview.Heading", font=(_UI_FONT, 10, "bold"))
    style.configure("Elenco.Treeview", font=(_UI_FONT, 10), rowheight=28)
    style.configure("Elenco.Treeview.Heading", font=(_UI_FONT, 10, "bold"))
    if sys.platform == "darwin":
        style.map("Treeview", background=[], foreground=[])


class _Suggerimento:
    """Fumetto con una frase che compare dopo mezzo secondo con il mouse fermo sul
    widget: spiega a cosa serve un pulsante senza occupare spazio."""

    RITARDO_MS = 600

    def __init__(self, widget, testo):
        self.widget, self.testo = widget, testo
        self._finestra = None
        self._after = None
        widget.bind("<Enter>", self._programma, add="+")
        widget.bind("<Leave>", self._nascondi, add="+")
        widget.bind("<ButtonPress>", self._nascondi, add="+")

    def _programma(self, _e=None):
        self._nascondi()
        self._after = self.widget.after(self.RITARDO_MS, self._mostra)

    def _mostra(self):
        self._after = None
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            self._finestra = tk.Toplevel(self.widget)
            self._finestra.wm_overrideredirect(True)
            self._finestra.wm_geometry(f"+{x}+{y}")
            tk.Label(self._finestra, text=self.testo, bg="#FFFFE1", fg="#1B1B1B",
                     relief=tk.SOLID, borderwidth=1, padx=6, pady=3, justify=tk.LEFT,
                     wraplength=320, font=(_UI_FONT, 9)).pack()
        except tk.TclError:
            self._finestra = None

    def _nascondi(self, _e=None):
        if self._after is not None:
            try:
                self.widget.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None
        if self._finestra is not None:
            try:
                self._finestra.destroy()
            except tk.TclError:
                pass
            self._finestra = None


class _ScrollableFrame(ttk.Frame):
    """Frame con scrollbar verticale (per la scheda Bollettino, che disegna
    una griglia di celle colorate e non un semplice elenco). Il contenuto
    segue la larghezza della finestra; la barra orizzontale compare solo se
    il contenuto non ci sta nemmeno alla sua larghezza minima."""

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        sfondo = _sfondo_tema()
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=sfondo)
        vscroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self._hscroll = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self._hscroll_visibile = False
        self.inner = tk.Frame(self.canvas, bg=sfondo)
        self.inner.bind("<Configure>", self._adatta)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>", self._adatta)
        self.canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=self._hscroll.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.bind("<Enter>", self._bind_wheel)
        self.canvas.bind("<Leave>", self._unbind_wheel)

    def _adatta(self, _e=None):
        larg_canvas = self.canvas.winfo_width()
        richiesta = self.inner.winfo_reqwidth()
        self.canvas.itemconfig(self._window, width=max(larg_canvas, richiesta))
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        serve = richiesta > larg_canvas > 1
        if serve and not self._hscroll_visibile:
            self._hscroll.pack(side=tk.BOTTOM, fill=tk.X, before=self.canvas)
            self._hscroll_visibile = True
        elif not serve and self._hscroll_visibile:
            self._hscroll.pack_forget()
            self._hscroll_visibile = False

    def _bind_wheel(self, _e):
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)
        self.canvas.bind_all("<Button-4>", self._on_wheel)
        self.canvas.bind_all("<Button-5>", self._on_wheel)

    def _unbind_wheel(self, _e):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_wheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class _BarraFluida(ttk.Frame):
    """Barra di pulsanti che va a capo da sola: i pulsanti si dispongono su
    quante righe servono in base alla larghezza disponibile, cosi' non
    obbligano la finestra a essere larga e non vengono mai tagliati.
    Usa ``place`` (non ``grid``): ogni riga e' indipendente, senza colonne
    condivise che allargherebbero tutte le righe al pulsante piu' largo."""

    _SPAZIO = 4

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._widget = []
        self._firma = None
        self.bind("<Configure>", self._riordina)

    def aggiungi(self, widget):
        """``widget`` deve avere questa barra come genitore."""
        self._widget.append(widget)
        self._firma = None
        self._riordina()

    def _riordina(self, _e=None):
        larghezza = self.winfo_width()
        if larghezza <= 1 or not self._widget:
            return
        x, y, alt_riga, posizioni = 0, self._SPAZIO, 0, []
        for w in self._widget:
            lw, lh = w.winfo_reqwidth(), w.winfo_reqheight()
            if x > 0 and x + lw > larghezza:
                x, y, alt_riga = 0, y + alt_riga + self._SPAZIO, 0
            posizioni.append((x, y))
            x += lw + self._SPAZIO
            alt_riga = max(alt_riga, lh)
        altezza = y + alt_riga + self._SPAZIO
        firma = (tuple(posizioni), altezza)
        if firma == self._firma:        # niente da spostare (evita cicli di <Configure>)
            return
        self._firma = firma
        for w, (px, py) in zip(self._widget, posizioni):
            w.place(x=px, y=py)
        self.configure(height=altezza)


class PollineCounterGUI:

    def __init__(self, root):
        self.root = root
        self.root.title("Conta Pollinica")
        _imposta_tema(root)
        self._imposta_dimensioni_iniziali()

        self.output_dir = percorsi.EXE_DIR
        self.settimana = None
        self.journal = None
        self.percorso_salvato = None
        self.soglie = {}
        self.text_log = None         # esistono solo nella schermata di lavoro
        self.lbl_stato = None
        self.lbl_file = None
        self._ora_salvataggio = None
        self._modificato_val = False
        self._undo_consecutivi = 0
        self._refresh_in_attesa = False
        self._tab_sporche = set()    # indici delle tab live da ridisegnare quando diventano visibili
        self._tree_cache = {}        # {Treeview: {iid: valori}} per aggiornare le righe sul posto
        self._giorno_bottoni = {}
        self._voce = None         # (Ascoltatore, Sintesi) quando la voce e' attiva
        self._voce_after = None
        self._talkback = False       # la voce ripete cio' che ha registrato
        self._voce_vocabolario = {}
        self._voce_attivazione = dominio.PAROLA_ATTIVAZIONE_DEFAULT
        # Interruttori (pulsanti e menu condividono la stessa variabile).
        self._var_beep = tk.BooleanVar(value=False)
        self._var_voce = tk.BooleanVar(value=False)
        self._var_talkback = tk.BooleanVar(value=False)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._collega_scorciatoie()
        self._carica_config_iniziale()
        self._build_avvio()

    # ================================================================
    # Stato "modifiche non salvate" (barra di stato in basso)
    # ================================================================
    @property
    def _modificato(self):
        return self._modificato_val

    @_modificato.setter
    def _modificato(self, valore):
        self._modificato_val = valore
        self._aggiorna_stato()

    def _aggiorna_stato(self):
        if self.lbl_stato is None or self.settimana is None:
            return
        if self._modificato_val:
            testo, colore = "● Modifiche non salvate", _COLORI["avviso"]
        elif self._ora_salvataggio:
            testo, colore = f"✓ Salvato alle {self._ora_salvataggio}", _COLORI["ok"]
        else:
            testo, colore = "Nessuna modifica da salvare", _COLORI["grigio"]
        try:
            self.lbl_stato.config(text=testo, foreground=colore)
            if self.percorso_salvato:
                self.lbl_file.config(text=f"File: {self.percorso_salvato.name}")
            else:
                self.lbl_file.config(
                    text=f"File: {self.settimana.nome_file_default()} (non ancora salvato)")
        except tk.TclError:             # barra gia' distrutta (cambio schermata)
            pass

    # ================================================================
    # Dimensioni della finestra
    # ================================================================
    def _imposta_dimensioni_iniziali(self):
        """Finestra proporzionata allo schermo (o all'ultima dimensione usata),
        centrata, con un minimo sotto il quale i pulsanti andrebbero a capo
        senza spazio."""
        layout = sessione.leggi_layout_finestra(percorsi.CONFIG_FILE)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        larg = min(layout.get("larghezza", min(int(sw * 0.9), 1400)), sw)
        alt = min(layout.get("altezza", min(int(sh * 0.85), 900)), sh)
        self._divisione = layout.get("divisione", 0.55)    # quota della larghezza per il pannello sinistro
        self.root.geometry(f"{larg}x{alt}+{(sw - larg) // 2}+{max(0, (sh - alt) // 2 - 20)}")
        self.root.minsize(min(760, sw), min(560, sh))

    def _salva_layout(self):
        """Ricorda dimensione e proporzione per la prossima volta (mai un errore
        all'utente: e' solo una comodita')."""
        try:
            if self.root.state() == "iconic":
                return
            sessione.salva_layout_finestra(percorsi.CONFIG_FILE, self.root.winfo_width(),
                                           self.root.winfo_height(), self._divisione)
        except (OSError, tk.TclError):
            pass

    # ================================================================
    # Configurazione cartella di lavoro
    # ================================================================
    def _carica_config_iniziale(self):
        anno = datetime.now().year
        cartella = sessione.leggi_cartella_anno(percorsi.CONFIG_FILE, anno)
        if cartella:
            self.output_dir = cartella
            return

        messagebox.showinfo(
            "Cartella di lavoro",
            f"Scegli la cartella dove salvare i file di questa stagione "
            f"({anno}).\nPotrai cambiarla in seguito dalla schermata iniziale.",
        )
        scelta = filedialog.askdirectory(
            title="Cartella di lavoro", initialdir=str(self.output_dir))
        cartella = Path(scelta) if scelta else self.output_dir
        sessione.salva_cartella_anno(percorsi.CONFIG_FILE, anno, cartella)
        self.output_dir = cartella

    def _cambia_cartella(self):
        scelta = filedialog.askdirectory(
            title="Cartella di lavoro", initialdir=str(self.output_dir))
        if not scelta:
            return
        self.output_dir = Path(scelta)
        sessione.salva_cartella_anno(percorsi.CONFIG_FILE, datetime.now().year, self.output_dir)
        self._aggiorna_lista_sessioni()

    def _apri_cartella(self):
        self._apri_con_sistema(self.output_dir)

    def _apri_con_sistema(self, percorso):
        """Apre un file o una cartella con il programma predefinito del sistema."""
        try:
            if sys.platform == "win32":
                os.startfile(str(percorso))
            else:
                comando = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([comando, str(percorso)],
                                 stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except (OSError, ValueError) as e:
            messagebox.showwarning("Non riesco ad aprirlo",
                                   f"Il sistema non ha aperto:\n{percorso}\n\n{e}")

    # ================================================================
    # Menu, scorciatoie e aiuto
    # ================================================================
    def _imposta_menu(self, lavoro):
        """Barra dei menu in stile Windows. Ogni voce chiama la stessa funzione
        del pulsante corrispondente: un solo percorso per ogni azione."""
        barra = tk.Menu(self.root)

        def menu(titolo):
            m = tk.Menu(barra, tearoff=False)
            barra.add_cascade(label=titolo, menu=m)
            return m

        def azione(lettera):
            return lambda: self._esegui_azione(lettera)

        f = menu("File")
        if lavoro:
            f.add_command(label="Salva", accelerator="Ctrl+S", command=azione("s"))
            f.add_command(label="Chiudi la settimana…", command=azione("q"))
        else:
            f.add_command(label="Inizia nuova settimana", command=self._avvio_nuovo_file)
            f.add_command(label="Apri un file da un'altra cartella…", command=self._avvio_importa)
            f.add_command(label="Cambia cartella di lavoro…", command=self._cambia_cartella)
        f.add_command(label="Apri la cartella di lavoro", command=self._apri_cartella)
        f.add_separator()
        f.add_command(label="Esci dal programma", accelerator="Ctrl+Q", command=self.on_closing)

        if lavoro:
            m = menu("Modifica")
            m.add_command(label="Annulla l'ultimo inserimento", accelerator="Ctrl+Z", command=azione("u"))
            m.add_command(label="Correggi un giorno…", command=azione("c"))
            m.add_command(label="Aggiungi una nota…", command=azione("n"))
            m.add_command(label="Fattore di conversione…", command=azione("f"))
            m.add_separator()
            m.add_command(label="Ultimi inserimenti", command=azione("l"))

            m = menu("Visualizza")
            for i, nome in enumerate(("Settimanale", "Giornaliero", "Bollettino", "Codici")):
                m.add_command(label=nome, accelerator=f"Ctrl+{i + 1}",
                              command=lambda i=i: self.notebook.select(i))

            m = menu("Strumenti")
            m.add_checkbutton(label="Beep a ogni inserimento", variable=self._var_beep,
                              command=azione("b"))
            m.add_checkbutton(label="Lettura vocale", variable=self._var_voce,
                              command=self._toggle_voce)
            m.add_checkbutton(label="Ripeti a voce", variable=self._var_talkback,
                              command=self._toggle_talkback)
            m.add_separator()
            m.add_command(label="Elenco codici stampabile (Word)…",
                          command=self._stampa_cheatsheet_codici)

        m = menu("Aiuto")
        m.add_command(label="Comandi e scorciatoie", accelerator="F1", command=self._mostra_aiuto)
        m.add_command(label="Guida completa…", command=self._apri_guida)
        m.add_separator()
        m.add_command(label="Informazioni", command=self._informazioni)
        self.root.config(menu=barra)
        self._menubar = barra

    def _collega_scorciatoie(self):
        """Scorciatoie classiche di Windows, valide in tutta la finestra principale
        (i dialoghi hanno i loro widget e non le ricevono)."""
        def lega(sequenza, comando, solo_lavoro=True):
            def gestore(_e=None):
                if self.settimana is not None or not solo_lavoro:
                    comando()
                return "break"
            self.root.bind(sequenza, gestore)
            if "Control-" in sequenza and sequenza[-2].islower():     # anche con Bloccamaiusc
                self.root.bind(sequenza[:-2] + sequenza[-2].upper() + ">", gestore)

        lega("<Control-s>", lambda: self._esegui_azione("s"))
        lega("<Control-z>", lambda: self._esegui_azione("u"))
        lega("<Control-q>", self.on_closing, solo_lavoro=False)
        lega("<F1>", self._mostra_aiuto, solo_lavoro=False)
        for i in range(4):
            lega(f"<Control-Key-{i + 1}>", lambda i=i: self.notebook.select(i))

    def _mostra_aiuto(self):
        if self.settimana is None:
            messagebox.showinfo("Come si comincia", HELP_AVVIO)
        else:
            messagebox.showinfo("Comandi", HELP_TEXT)

    def _apri_guida(self):
        """Apre il file di istruzioni per questo sistema, se si trova."""
        nome, sottocartella = {
            "win32": ("ISTRUZIONI_WINDOWS.txt", "windows"),
            "darwin": ("ISTRUZIONI_MAC.txt", "mac"),
        }.get(sys.platform, ("ISTRUZIONI_LINUX.txt", "script_aiuto"))
        for cartella in (percorsi.EXE_DIR, percorsi.EXE_DIR.parent / sottocartella,
                         percorsi.EXE_DIR.parent):
            for candidato in (cartella / nome, cartella / "ISTRUZIONI.txt"):
                if candidato.is_file():
                    self._apri_con_sistema(candidato)
                    return
        messagebox.showinfo("Guida non trovata",
                            f"Non trovo il file {nome} accanto al programma.\n"
                            "Premi F1 per la guida rapida dei comandi.")

    def _informazioni(self):
        messagebox.showinfo(
            "Informazioni",
            "Conta Pollinica\n\n"
            "Registra i granuli di polline contati al microscopio, salva i dati in Excel e "
            "prepara il bollettino settimanale.\n\n"
            f"Cartella di lavoro:\n{self.output_dir}")

    # ================================================================
    # Schermata iniziale: scelta sessione/file/nuovo
    # ================================================================
    @staticmethod
    def _abbrevia_percorso(percorso, massimo=42):
        testo = str(percorso)
        return testo if len(testo) <= massimo else "…" + testo[-(massimo - 1):]

    def _build_avvio(self):
        self._imposta_menu(lavoro=False)
        self.root.title("Conta Pollinica")
        self.frame_avvio = ttk.Frame(self.root, padding=(28, 16))
        self.frame_avvio.pack(fill=tk.BOTH, expand=True)

        # Le barre in basso si impacchettano PRIMA di quanto si espande: con la
        # finestra bassa e' l'elenco a restringersi, non i pulsanti.
        in_basso = ttk.Frame(self.frame_avvio)
        in_basso.pack(side=tk.BOTTOM, fill=tk.X, pady=(14, 0))
        riga_azioni = ttk.Frame(in_basso)
        riga_azioni.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        ttk.Button(riga_azioni, text="Esci", width=10, command=self.on_closing).pack(side=tk.RIGHT)
        ttk.Button(riga_azioni, text="Guida (F1)", command=self._mostra_aiuto).pack(side=tk.RIGHT, padx=8)
        ttk.Button(riga_azioni, text="Elenco codici stampabile",
                   command=self._stampa_cheatsheet_codici).pack(side=tk.LEFT)

        riga_cartella = ttk.Frame(in_basso)
        riga_cartella.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(riga_cartella, text="Apri", command=self._apri_cartella).pack(side=tk.RIGHT)
        ttk.Button(riga_cartella, text="Cambia…", command=self._cambia_cartella).pack(side=tk.RIGHT, padx=8)
        ttk.Label(riga_cartella, text="Cartella di lavoro:", style="Sezione.TLabel").pack(side=tk.LEFT)
        self.lbl_cartella = ttk.Label(riga_cartella, text="", anchor="w")
        self.lbl_cartella.pack(side=tk.LEFT, padx=8)
        ttk.Separator(in_basso).pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 10))

        # Intestazione
        oggi = datetime.now()
        ttk.Label(self.frame_avvio, text="Conta Pollinica", style="Titolo.TLabel").pack(anchor="w")
        ttk.Label(self.frame_avvio, foreground=_COLORI["grigio"],
                  text=f"Stagione {oggi.year}  ·  oggi è {dominio.nome_giorno_accentato(oggi)} "
                       f"{oggi.strftime('%d/%m/%Y')}").pack(anchor="w", pady=(0, 8))

        # Avviso di lavoro interrotto (riempito da _aggiorna_lista_sessioni)
        self.box_avviso = ttk.Frame(self.frame_avvio)
        self.box_avviso.pack(side=tk.TOP, fill=tk.X)

        # Nuova settimana
        nuova = ttk.LabelFrame(self.frame_avvio, text=" Nuova settimana ", padding=(14, 6, 14, 8))
        nuova.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        riga_data = ttk.Frame(nuova)
        riga_data.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(riga_data, text="Un giorno della settimana:").pack(side=tk.LEFT)
        ttk.Button(riga_data, text="◀", width=3,
                   command=lambda: self._sposta_settimana(-7)).pack(side=tk.LEFT, padx=(10, 2))
        self._var_data = tk.StringVar(value=dominio.lunedi_di(oggi).strftime("%d-%m-%Y"))
        self.entry_settimana = ttk.Entry(riga_data, width=14, textvariable=self._var_data,
                                         font=(_MONO_FONT, 11))
        self.entry_settimana.pack(side=tk.LEFT)
        ttk.Button(riga_data, text="▶", width=3,
                   command=lambda: self._sposta_settimana(7)).pack(side=tk.LEFT, padx=(2, 12))
        self.btn_nuova = ttk.Button(riga_data, text="Inizia nuova settimana", style="Accent.TButton",
                                    command=self._avvio_nuovo_file)
        self.btn_nuova.pack(side=tk.RIGHT)
        self.lbl_anteprima = ttk.Label(nuova, text="", foreground=_COLORI["grigio"])
        self.lbl_anteprima.pack(side=tk.TOP, anchor="w", pady=(6, 0))
        self._var_data.trace_add("write", lambda *_: self._aggiorna_anteprima_settimana())
        self.entry_settimana.bind("<Return>", lambda e: self._avvio_nuovo_file())
        self._aggiorna_anteprima_settimana()

        # Riprendere un lavoro
        riprendi = ttk.LabelFrame(self.frame_avvio, text=" Riprendi un lavoro ", padding=(14, 6, 14, 8))
        riprendi.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        riga_bottoni = ttk.Frame(riprendi)
        riga_bottoni.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        self.btn_riprendi = ttk.Button(riga_bottoni, text="Riprendi selezionato", style="Accent.TButton",
                                       command=self._avvio_riprendi_selezionata, state=tk.DISABLED)
        self.btn_riprendi.pack(side=tk.LEFT)
        self.btn_rimuovi = ttk.Button(riga_bottoni, text="Rimuovi non salvato",
                                      command=self._avvio_rimuovi_selezionata, state=tk.DISABLED)
        self.btn_rimuovi.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(riga_bottoni, text="Apri un file da un'altra cartella…",
                   command=self._avvio_importa).pack(side=tk.LEFT, padx=8)
        self.lbl_elenco_info = ttk.Label(riprendi, text="", foreground=_COLORI["grigio"])
        self.lbl_elenco_info.pack(side=tk.TOP, anchor="w", pady=(0, 4))

        elenco = ttk.Frame(riprendi)
        elenco.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        cols = ("stato", "nome", "quando", "info")
        self.tree_avvio = ttk.Treeview(elenco, columns=cols, show="headings", height=4,
                                       style="Elenco.Treeview", selectmode="browse")
        for col, titolo, larg, elastica in (("stato", "Stato", 100, False),
                                            ("nome", "Settimana / file", 300, True),
                                            ("quando", "Ultima modifica", 120, False),
                                            ("info", "Dettagli", 190, True)):
            self.tree_avvio.heading(col, text=titolo, anchor=tk.W)
            self.tree_avvio.column(col, width=larg, minwidth=80, stretch=elastica, anchor=tk.W)
        scroll = ttk.Scrollbar(elenco, command=self.tree_avvio.yview)
        self.tree_avvio.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_avvio.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_avvio.bind("<Double-1>", lambda e: self._avvio_riprendi_selezionata())
        self.tree_avvio.bind("<<TreeviewSelect>>", lambda e: self._aggiorna_pulsanti_avvio())

        self._aggiorna_lista_sessioni()
        self.entry_settimana.focus_set()

    # ── data della nuova settimana ──
    def _data_settimana_scelta(self):
        """datetime scelto nella casella (oggi se vuota), o None se non valido."""
        testo = self._var_data.get().strip()
        return dominio.parse_data_flessibile(testo) if testo else datetime.now()

    def _aggiorna_anteprima_settimana(self):
        dt = self._data_settimana_scelta()
        if dt is None:
            self.lbl_anteprima.config(
                text="Non riconosco questa data. Prova con: 9-2-2026, 9/2/2026, 9 feb 2026.",
                foreground=_COLORI["errore"])
            self.btn_nuova.config(state=tk.DISABLED)
        else:
            self.lbl_anteprima.config(
                text=f"Settimana {dominio.descrivi_settimana(dominio.lunedi_di(dt))}",
                foreground=_COLORI["grigio"])
            self.btn_nuova.config(state=tk.NORMAL)

    def _sposta_settimana(self, giorni):
        dt = self._data_settimana_scelta() or datetime.now()
        self._var_data.set((dominio.lunedi_di(dt) + timedelta(days=giorni)).strftime("%d-%m-%Y"))

    def _aggiorna_lista_sessioni(self):
        self.lbl_cartella.config(text=self._abbrevia_percorso(self.output_dir))
        for item in self.tree_avvio.get_children():
            self.tree_avvio.delete(item)
        for figlio in self.box_avviso.winfo_children():
            figlio.destroy()
        self._righe_avvio = []
        adesso = datetime.now()

        def quando(path):
            try:
                return dominio.descrivi_quando(datetime.fromtimestamp(path.stat().st_mtime), adesso)
            except OSError:
                return ""

        interrotte = sessione.recupera_sessioni(self.output_dir)
        for path, info in interrotte:
            self.tree_avvio.insert("", tk.END, values=(
                "Non salvato", self._descrivi_journal(info).capitalize(), quando(path),
                f"{info['n_operazioni']} operazioni da recuperare",
            ))
            self._righe_avvio.append(("journal", path, info))

        for f in sessione.cerca_file_ripresa(self.output_dir, esportatori.TEMPLATE_FILE.name):
            self.tree_avvio.insert("", tk.END, values=(
                "File salvato", f.name, quando(f), sessione.conta_righe_log(f),
            ))
            self._righe_avvio.append(("xlsx", f, None))

        self._aggiorna_pulsanti_avvio()
        self.lbl_elenco_info.config(
            text=("Doppio clic su una riga per aprirla." if self._righe_avvio else
                  "Nessun lavoro precedente in questa cartella: inizia una nuova settimana qui sopra."))
        if interrotte:
            self._mostra_avviso_interrotte(interrotte)

    @staticmethod
    def _descrivi_journal(info):
        try:
            return dominio.descrivi_settimana(datetime.strptime(info["lunedi"], "%Y-%m-%d"))
        except ValueError:
            return f"settimana del {info['lunedi']}"

    def _aggiorna_pulsanti_avvio(self):
        """Riprendi: una riga qualsiasi. Rimuovi: solo righe 'Non salvato' (mai i
        file .xlsx salvati)."""
        sel = self.tree_avvio.selection()
        self.btn_riprendi.config(state=(tk.NORMAL if sel else tk.DISABLED))
        solo_journal = bool(sel) and all(
            self._righe_avvio[self.tree_avvio.index(i)][0] == "journal" for i in sel)
        self.btn_rimuovi.config(state=(tk.NORMAL if solo_journal else tk.DISABLED))

    def _mostra_avviso_interrotte(self, interrotte):
        """Riquadro giallo in cima: dopo un blackout o una chiusura improvvisa e'
        la prima cosa che l'operatore deve vedere. 'interrotte' e' l'elenco di
        sessione.recupera_sessioni (la piu' recente per prima)."""
        path, info = interrotte[0]
        settimana = self._descrivi_journal(info)
        if len(interrotte) == 1:
            testo = (f"C'è un lavoro non salvato della settimana {settimana} ({info['n_operazioni']} operazioni). "
                     "I dati sono al sicuro e si possono riprendere, oppure rimuovere "
                     "se non servono più.")
        else:
            testo = (f"Ci sono {len(interrotte)} lavori non salvati; il più recente è della settimana {settimana} "
                     f"({info['n_operazioni']} operazioni). I dati sono al sicuro: si può riprendere "
                     "il più recente, sceglierne un altro oppure rimuovere quelli che non servono più.")
        riquadro = tk.Frame(self.box_avviso, bg=_COLORI["riquadro_avviso"],
                            highlightbackground="#E8D38A", highlightthickness=1)
        riquadro.pack(fill=tk.X, pady=(0, 8))
        bottoni = tk.Frame(riquadro, bg=_COLORI["riquadro_avviso"])
        bottoni.pack(side=tk.RIGHT, padx=10, pady=8)
        ttk.Button(bottoni, text="Riprendi" + (" il più recente" if len(interrotte) > 1 else ""),
                   style="Accent.TButton",
                   command=lambda: self._avvia_da_journal(path)).pack(side=tk.TOP, fill=tk.X)
        ttk.Button(bottoni, text="Scegli o rimuovi…",
                   command=self._dialog_sessioni_interrotte).pack(side=tk.TOP, fill=tk.X, pady=(6, 0))
        tk.Label(riquadro, text="⚠  " + testo, bg=_COLORI["riquadro_avviso"], fg="#1B1B1B",
                 font=(_UI_FONT, 10), justify=tk.LEFT, anchor="w", wraplength=520,
                 padx=12, pady=8).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _dialog_sessioni_interrotte(self):
        """Elenco dei soli lavori non salvati: riprenderne uno, rimuoverne alcuni
        (selezione multipla) o tutti."""
        top = tk.Toplevel(self.root)
        top.title("Lavori non salvati")
        top.minsize(560, 300)

        ttk.Label(top, text="Seleziona una o più righe (Ctrl o Maiusc + clic).",
                  foreground=_COLORI["grigio"]).pack(side=tk.TOP, anchor="w", padx=14, pady=(12, 6))
        riga_bottoni = ttk.Frame(top)
        riga_bottoni.pack(side=tk.BOTTOM, fill=tk.X, padx=14, pady=12)
        elenco = ttk.Frame(top)
        elenco.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=14)
        cols = ("nome", "quando", "info")
        tree = ttk.Treeview(elenco, columns=cols, show="headings", height=6,
                            style="Elenco.Treeview", selectmode="extended")
        for col, titolo, larg, elastica in (("nome", "Settimana", 260, True),
                                            ("quando", "Ultima modifica", 120, False),
                                            ("info", "Operazioni", 100, False)):
            tree.heading(col, text=titolo, anchor=tk.W)
            tree.column(col, width=larg, minwidth=80, stretch=elastica, anchor=tk.W)
        scroll = ttk.Scrollbar(elenco, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        voci = []   # [(path, info)] nello stesso ordine delle righe

        def ricarica():
            voci[:] = sessione.recupera_sessioni(self.output_dir)
            for item in tree.get_children():
                tree.delete(item)
            adesso = datetime.now()
            for path, info in voci:
                try:
                    q = dominio.descrivi_quando(datetime.fromtimestamp(path.stat().st_mtime), adesso)
                except OSError:
                    q = ""
                tree.insert("", tk.END, values=(self._descrivi_journal(info).capitalize(),
                                                q, info["n_operazioni"]))
            aggiorna_pulsanti()

        def selezionate():
            return [voci[tree.index(i)] for i in tree.selection()]

        def aggiorna_pulsanti(*_):
            n = len(tree.selection())
            btn_riprendi.config(state=(tk.NORMAL if n == 1 else tk.DISABLED))
            btn_sel.config(state=(tk.NORMAL if n >= 1 else tk.DISABLED))
            btn_tutte.config(state=(tk.NORMAL if voci else tk.DISABLED))

        def riprendi(*_):
            scelte = selezionate()
            if len(scelte) != 1:
                return
            top.destroy()
            self._avvia_da_journal(scelte[0][0])

        def rimuovi(da_rimuovere):
            if da_rimuovere and self._rimuovi_sessioni_interrotte(da_rimuovere, parent=top):
                ricarica()
                if not voci:
                    top.destroy()

        btn_riprendi = ttk.Button(riga_bottoni, text="Riprendi", style="Accent.TButton", command=riprendi)
        btn_riprendi.pack(side=tk.LEFT)
        btn_sel = ttk.Button(riga_bottoni, text="Rimuovi selezionate",
                             command=lambda: rimuovi(selezionate()))
        btn_sel.pack(side=tk.LEFT, padx=8)
        btn_tutte = ttk.Button(riga_bottoni, text="Rimuovi tutte", command=lambda: rimuovi(list(voci)))
        btn_tutte.pack(side=tk.LEFT)
        ttk.Button(riga_bottoni, text="Chiudi", command=top.destroy).pack(side=tk.RIGHT)
        tree.bind("<<TreeviewSelect>>", aggiorna_pulsanti)
        tree.bind("<Double-1>", riprendi)
        ricarica()
        self._prepara_dialogo(top)

    def _rimuovi_sessioni_interrotte(self, voci, parent=None):
        """Cancella definitivamente i journal di 'voci' [(path, info)] dopo conferma.
        Un solo percorso per dialogo ed elenco. Ritorna True se ha rimosso qualcosa."""
        parent = parent or self.root
        n_op = sum(info["n_operazioni"] for _, info in voci)
        if len(voci) == 1:
            domanda = (f"Rimuovere il lavoro non salvato della settimana {self._descrivi_journal(voci[0][1])} "
                       f"({n_op} operazioni)?")
        else:
            domanda = f"Rimuovere {len(voci)} lavori non salvati ({n_op} operazioni in tutto)?"
        if not messagebox.askyesno(
                "Rimuovere lavoro non salvato",
                domanda + "\n\nI conteggi non salvati andranno persi definitivamente: "
                          "non si può tornare indietro.",
                icon="warning", default="no", parent=parent):
            return False
        falliti = sessione.rimuovi_sessioni([p for p, _ in voci])
        if falliti:
            elenco = "\n".join(f"  {p.name}" for p, _ in falliti)
            messagebox.showerror(
                "Rimozione non riuscita",
                f"Non è stato possibile rimuovere:\n{elenco}\n\n"
                "Controlla che il file non sia aperto in un altro programma e riprova.",
                parent=parent)
        self._aggiorna_lista_sessioni()
        return len(falliti) < len(voci)

    def _avvio_rimuovi_selezionata(self):
        voci = []
        for item in self.tree_avvio.selection():
            tipo, path, info = self._righe_avvio[self.tree_avvio.index(item)]
            if tipo == "journal":
                voci.append((path, info))
        if voci:
            self._rimuovi_sessioni_interrotte(voci)

    def _avvio_riprendi_selezionata(self):
        sel = self.tree_avvio.selection()
        if not sel:
            messagebox.showwarning("Nessuna selezione", "Seleziona una sessione o un file dall'elenco.")
            return
        tipo, path, _ = self._righe_avvio[self.tree_avvio.index(sel[0])]
        if tipo == "journal":
            self._avvia_da_journal(path)
        else:
            self._avvia_da_xlsx(path)

    def _avvio_importa(self):
        raw = filedialog.askopenfilename(
            title="Importa file conta pollinica", initialdir=str(self.output_dir),
            filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
        )
        if not raw:
            return
        self._avvia_da_xlsx(Path(raw))

    def _avvio_nuovo_file(self):
        dt = self._data_settimana_scelta()
        if dt is None:
            messagebox.showerror("Data non valida",
                                 "Non riconosco questa data. Prova con: 9-2-2026, 9/2/2026, 9 feb 2026.")
            return
        lunedi = dominio.lunedi_di(dt)

        # Iniziare "da zero" una settimana gia' iniziata cancellerebbe il lavoro non
        # salvato (Journal.avvia sovrascrive): si offre prima di riprenderlo.
        percorso_journal = self.output_dir / sessione.nome_journal(lunedi)
        if percorso_journal.exists():
            risp = messagebox.askyesnocancel(
                "Settimana già iniziata",
                "Per questa settimana c'è già un lavoro non salvato.\n\n"
                "Sì: riprendi quel lavoro\n"
                "No: ricomincia da zero (il lavoro non salvato verrà cancellato)\n"
                "Annulla: torna indietro")
            if risp is None:
                return
            if risp:
                self._avvia_da_journal(percorso_journal)
                return
        else:
            nome = sessione.Settimana(lunedi).nome_file_default()
            if (self.output_dir / nome).exists():
                risp = messagebox.askyesnocancel(
                    "File già presente",
                    f"Nella cartella di lavoro c'è già il file:\n{nome}\nper questa settimana.\n\n"
                    "Sì: aprilo e continua da lì\n"
                    "No: crea comunque una settimana nuova\n"
                    "Annulla: torna indietro")
                if risp is None:
                    return
                if risp:
                    self._avvia_da_xlsx(self.output_dir / nome)
                    return

        self.settimana = sessione.Settimana(lunedi)
        self.journal = sessione.Journal(percorso_journal)
        self.journal.avvia(self.settimana)
        self.percorso_salvato = None
        self._entra_in_lavoro()

    def _avvia_da_journal(self, path):
        try:
            self.settimana = sessione.ripristina_da_journal(path)
        except ValueError as e:
            messagebox.showerror("Sessione danneggiata", str(e))
            return
        self.journal = sessione.Journal.riprendi(path)
        self.percorso_salvato = None
        self._entra_in_lavoro()

    def _avvia_da_xlsx(self, path):
        try:
            self.settimana = sessione.carica_da_xlsx(path)
        except ValueError as e:
            messagebox.showerror("File non valido", str(e))
            return

        lunedi_attuale = self.settimana.lunedi
        domenica = lunedi_attuale + timedelta(days=6)
        mantieni = messagebox.askyesno(
            "Settimana",
            f"Il file si riferisce alla settimana dal "
            f"{lunedi_attuale.strftime('%d-%m-%Y')} al {domenica.strftime('%d-%m-%Y')}.\n\n"
            f"Mantenere questa settimana?",
        )
        if not mantieni:
            nuova = self._chiedi_settimana_dialogo(lunedi_attuale)
            if nuova is None:
                self.settimana = None
                return
            self.settimana.lunedi = nuova

        # Journal.avvia sovrascrive: non farlo in silenzio se c'e' un lavoro interrotto
        if (self.output_dir / sessione.nome_journal(self.settimana.lunedi)).exists():
            if not messagebox.askyesno(
                    "Lavoro non salvato",
                    "Per questa settimana c'è già un lavoro interrotto, non salvato.\n"
                    "Se continui verrà cancellato.\n\nContinuare?",
                    icon="warning", default="no"):
                self.settimana = None
                return

        self.percorso_salvato = None
        self.journal = sessione.Journal(self.output_dir / sessione.nome_journal(self.settimana.lunedi))
        self.journal.avvia(self.settimana)
        self._entra_in_lavoro()

    def _chiedi_settimana_dialogo(self, default_dt):
        testo = simpledialog.askstring(
            "Settimana",
            "Data della settimana (es. 9-2-2026, 9/2/2026, 9 feb 2026):",
            initialvalue=default_dt.strftime("%d-%m-%Y"), parent=self.root,
        )
        if not testo:
            return None
        dt = dominio.parse_data_flessibile(testo)
        if not dt:
            messagebox.showerror("Data non valida", "Data non riconosciuta.")
            return None
        return dominio.lunedi_di(dt)

    # ================================================================
    # Schermata di lavoro
    # ================================================================
    def _entra_in_lavoro(self):
        self.frame_avvio.destroy()
        self.soglie = esportatori.carica_soglie(self.output_dir) or {}
        self._ora_salvataggio = None
        self._build_lavoro()
        self._var_beep.set(self.settimana.beep)
        self._modificato = False
        self.settimana.on_change(self._on_settimana_change)

        if self.settimana.giorno_attivo:
            self._imposta_giorno_ui(self.settimana.giorno_attivo)
        else:
            self._log("Seleziona un giorno per iniziare l'inserimento.")
            self._imposta_entry_abilitata(False)

        self._tab_sporche = set(self._TAB_LIVE)
        self._tree_cache = {}
        self._refresh_tabs()

    def _build_lavoro(self):
        self._imposta_menu(lavoro=True)
        self.root.title(f"Conta Pollinica — settimana {dominio.descrivi_settimana(self.settimana.lunedi)}")
        self.frame_lavoro = ttk.Frame(self.root)
        self.frame_lavoro.pack(fill=tk.BOTH, expand=True)

        # Barra di stato in basso: impacchettata per prima, non si taglia mai.
        barra_stato = ttk.Frame(self.frame_lavoro, padding=(12, 4))
        barra_stato.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Separator(self.frame_lavoro).pack(side=tk.BOTTOM, fill=tk.X)
        self.lbl_file = ttk.Label(barra_stato, text="", foreground=_COLORI["grigio"])
        self.lbl_file.pack(side=tk.LEFT)
        self.lbl_stato = ttk.Label(barra_stato, text="")
        self.lbl_stato.pack(side=tk.RIGHT)

        self.pane = tk.PanedWindow(self.frame_lavoro, orient=tk.HORIZONTAL, sashwidth=6,
                                   bg="#d0d0d0", borderwidth=0)
        self.pane.pack(fill=tk.BOTH, expand=True)

        # ── Pannello sinistro ──
        sinistra = ttk.Frame(self.pane, padding=(10, 8, 6, 8))
        self.pane.add(sinistra, stretch="always", minsize=_MIN_SINISTRA)

        barra_giorni = ttk.Frame(sinistra)
        barra_giorni.pack(side=tk.TOP, fill=tk.X)
        for g in range(1, 8):
            nome = dominio.GIORNI_ABBREV[g - 1]
            btn = ttk.Button(barra_giorni, text=nome, style="Giorno.TButton",
                             command=lambda g=g: self._pick_giorno(g))
            btn.pack(side=tk.LEFT, padx=1, fill=tk.X, expand=True)
            self._giorno_bottoni[g] = btn
            _Suggerimento(btn, f"Inizia a contare {dominio.GIORNI_NOMI[g].upper()}")

        self.lbl_giorno = ttk.Label(sinistra, text="Nessun giorno selezionato", style="Giorno.TLabel",
                                    anchor="w")
        self.lbl_giorno.pack(side=tk.TOP, fill=tk.X, pady=(10, 6))

        # Le barre in basso si impacchettano PRIMA del riquadro di log, che si
        # espande: se la finestra e' bassa e' il log a restringersi, mai i pulsanti.
        ttk.Label(sinistra, foreground=_COLORI["grigio"], wraplength=320, justify=tk.LEFT,
                  text="Scrivi il numero del polline (es. 24) e premi Invio.  48x4 = quattro granuli "
                       "del codice 48.  Un punto (.) ripete l'ultimo."
                  ).pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        input_frame = ttk.Frame(sinistra, padding=(0, 8, 0, 0))
        input_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(input_frame, text="Codice:", style="Sezione.TLabel").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(input_frame, text="Invia", command=self._invia).pack(side=tk.RIGHT, padx=(8, 0))
        self.entry = ttk.Entry(input_frame, font=(_MONO_FONT, 12))
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.entry.bind("<Return>", self._invia)
        self.entry.focus_set()

        # I pulsanti vanno a capo da soli in base alla larghezza (_BarraFluida);
        # le etichette cambiano solo nei Checkbutton (interruttori), che hanno larghezza fissa.
        barra_azioni = _BarraFluida(sinistra)
        barra_azioni.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        barra_interruttori = _BarraFluida(sinistra)
        barra_interruttori.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

        def pulsante(testo, comando, suggerimento, **kw):
            b = ttk.Button(barra_azioni, text=testo, command=comando, **kw)
            barra_azioni.aggiungi(b)
            _Suggerimento(b, suggerimento)
            return b

        def interruttore(testo, variabile, comando, suggerimento):
            c = ttk.Checkbutton(barra_interruttori, text=testo, variable=variabile,
                                command=comando, style="Switch.TCheckbutton")
            barra_interruttori.aggiungi(c)
            _Suggerimento(c, suggerimento)
            return c

        pulsante("Aiuto (h)", lambda: self._esegui_azione("h"),
                 "Elenco dei comandi e delle scorciatoie — anche F1")
        pulsante("Correggi (c)", lambda: self._esegui_azione("c"),
                 "Cambia il valore di una specie in un giorno già contato")
        pulsante("Nota (n)", lambda: self._esegui_azione("n"),
                 "Aggiunge una nota di testo alla giornata (finisce nel file Excel)")
        pulsante("Annulla (u)", lambda: self._esegui_azione("u"),
                 "Annulla l'ultimo inserimento — anche Ctrl+Z")
        pulsante("Salva (s)", lambda: self._esegui_azione("s"),
                 "Salva il file Excel senza uscire — anche Ctrl+S", style="Accent.TButton")
        pulsante("Chiudi giornata (d)", lambda: self._esegui_azione("d"),
                 "Finisce la giornata corrente: poi scegli un altro giorno")
        pulsante("Chiudi settimana (q)", lambda: self._esegui_azione("q"),
                 "Salva (se vuoi) e torna alla schermata iniziale. Per uscire dal programma: Ctrl+Q")

        interruttore("Beep (b)", self._var_beep, lambda: self._esegui_azione("b"),
                     "Un suono a ogni granulo registrato")
        interruttore("Voce", self._var_voce, self._toggle_voce,
                     "Detta i granuli al microfono: ogni frase inizia con «conta»")
        interruttore("Ripeti a voce", self._var_talkback, self._toggle_talkback,
                     "Il computer ripete il nome di ciò che ha registrato")

        self.text_log = tk.Text(
            sinistra, wrap=tk.WORD, font=(_UI_FONT, 11), height=4, width=20,
            bg=_COLORI["sfondo_registro"], fg=_COLORI["testo_registro"],
            state=tk.DISABLED, relief=tk.FLAT, padx=8, pady=6,
            highlightthickness=1, highlightbackground="#D0D0D0", highlightcolor="#D0D0D0",
        )
        for tag, colore in (("inserimento", _COLORI["ok"]), ("annullo", _COLORI["avviso"]),
                            ("errore", _COLORI["errore"]), ("voce", _COLORI["accento"]),
                            ("ok", _COLORI["ok"])):
            self.text_log.tag_configure(tag, foreground=colore)
        self.text_log.tag_configure("grigio", foreground=_COLORI["grigio"])
        scroll_log = ttk.Scrollbar(sinistra, command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=scroll_log.set)
        scroll_log.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_log.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # ── Pannello destro ──
        destra = ttk.Frame(self.pane, padding=(6, 8, 10, 8))
        self.pane.add(destra, stretch="always", minsize=_MIN_DESTRA)

        # Il divisorio tiene la proporzione quando la finestra cambia misura;
        # trascinandolo a mano la nuova proporzione viene ricordata.
        self._divisione_in_corso = False
        self.pane.bind("<Configure>", self._applica_divisione)
        self.pane.bind("<ButtonRelease-1>", self._memorizza_divisione)

        self.notebook = ttk.Notebook(destra)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self._font_tab = tkfont.Font(family=_UI_FONT, size=10)

        self._build_tab_settimanale()
        self._build_tab_giornaliero()
        self._build_tab_bollettino()
        self._build_tab_codici()
        # Solo la tab visibile si ridisegna a ogni modifica: le altre, quando le si apre.
        self.notebook.bind("<<NotebookTabChanged>>", lambda _e: self._aggiorna_tab_visibile())

    def _applica_divisione(self, _e=None):
        if self._divisione_in_corso:
            return
        self._divisione_in_corso = True
        self.root.after_idle(self._posiziona_divisore)

    def _posiziona_divisore(self):
        self._divisione_in_corso = False
        try:
            larghezza = self.pane.winfo_width()
            if larghezza > 1:
                # la proporzione scelta, ma senza scendere sotto il minimo dei pannelli
                x = int(larghezza * self._divisione)
                x = min(x, larghezza - _MIN_DESTRA - 6)
                self.pane.sash_place(0, max(x, _MIN_SINISTRA), 0)
        except tk.TclError:             # finestra in chiusura o non ancora disegnata
            pass

    def _memorizza_divisione(self, _e=None):
        try:
            larghezza = self.pane.winfo_width()
            if larghezza > 1:
                self._divisione = min(sessione.DIVISIONE_FINESTRA_MAX,
                                      max(sessione.DIVISIONE_FINESTRA_MIN,
                                          self.pane.sash_coord(0)[0] / larghezza))
        except tk.TclError:
            pass

    def _crea_tabella(self, tab, intestazioni):
        """Treeview con scrollbar per una scheda. ``intestazioni`` e'
        {colonna: titolo}; le colonne di testo ("specie", "voce") sono a
        sinistra, le altre centrate. Il tag "famiglia" evidenzia i totali."""
        tree = ttk.Treeview(tab, columns=tuple(intestazioni), show="headings",
                            style="Summary.Treeview", height=4)
        for col, titolo in intestazioni.items():
            tree.heading(col, text=titolo)
            tree.column(col, anchor=tk.W if col in ("specie", "voce") else tk.CENTER)
        tree.tag_configure("famiglia", font=(_UI_FONT, 10, "bold"),
                           background=_COLORI["famiglia"], foreground="#000000")
        scroll = ttk.Scrollbar(tab, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        return tree

    def _colonne_tabella(self, tree, misure, elastiche, margine=18, minimo=90):
        """Colonne che seguono la scheda: quelle in ``misure`` ({colonna: testo
        piu' largo}) restano strette quanto serve, quelle in ``elastiche``
        dividono lo spazio che avanza (mai sotto ``minimo`` px)."""
        for col, testo in misure.items():
            larghezza = self._font_tab.measure(testo) + margine
            tree.column(col, width=larghezza, minwidth=larghezza, stretch=False)
        for col in elastiche:
            tree.column(col, width=minimo, minwidth=minimo, stretch=True)

    def _build_tab_settimanale(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Settimanale")

        # Totali in basso impacchettati per primi: non si tagliano mai.
        totals = ttk.Frame(tab, padding=(0, 8))
        totals.pack(side=tk.BOTTOM, fill=tk.X)
        self.lbl_s_pollini = ttk.Label(totals, text="Pollini: 0", font=(_UI_FONT, 11), anchor=tk.W)
        self.lbl_s_pollini.pack(fill=tk.X, padx=10)
        self.lbl_s_spore = ttk.Label(totals, text="Spore: 0", font=(_UI_FONT, 11), anchor=tk.W)
        self.lbl_s_spore.pack(fill=tk.X, padx=10)
        ttk.Separator(totals, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=4)
        self.lbl_s_totale = ttk.Label(totals, text="TOTALE: 0", font=(_UI_FONT, 12, "bold"), anchor=tk.W)
        self.lbl_s_totale.pack(fill=tk.X, padx=10)

        self.tree_sett = self._crea_tabella(
            tab, {"codice": "Cod.", "specie": "Specie", "conteggio": "Tot."})
        self._colonne_tabella(self.tree_sett, {"codice": "00000", "conteggio": "000000"}, ["specie"])

    def _build_tab_giornaliero(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Giornaliero")

        totals = ttk.Frame(tab, padding=(0, 8))
        totals.pack(side=tk.BOTTOM, fill=tk.X)
        # Qui i totali sono allineati in colonna ("LUN:3  MAR:5"): servono i caratteri a larghezza fissa.
        self.lbl_g_pollini = ttk.Label(totals, text="Pollini: -", font=(_MONO_FONT, 10), anchor=tk.W)
        self.lbl_g_pollini.pack(fill=tk.X, padx=10)
        self.lbl_g_spore = ttk.Label(totals, text="Spore: -", font=(_MONO_FONT, 10), anchor=tk.W)
        self.lbl_g_spore.pack(fill=tk.X, padx=10)
        ttk.Separator(totals, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=4)
        self.lbl_g_totale = ttk.Label(totals, text="TOTALE: -", font=(_MONO_FONT, 11, "bold"), anchor=tk.W)
        self.lbl_g_totale.pack(fill=tk.X, padx=10)

        giorni = {g.lower(): g for g in dominio.GIORNI_ABBREV}
        self.tree_giorn = self._crea_tabella(
            tab, {"codice": "Cod.", "specie": "Specie", **giorni})
        # nove colonne: margini stretti, cosi' alla finestra minima ci stanno tutte
        self._colonne_tabella(self.tree_giorn, dict.fromkeys(["codice", *giorni], "000"),
                              ["specie"], margine=12, minimo=60)

    def _build_tab_bollettino(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Bollettino")

        info_frame = ttk.Frame(tab, padding=(0, 6))
        info_frame.pack(side=tk.BOTTOM, fill=tk.X)
        riga_info = ttk.Frame(info_frame)
        riga_info.pack(fill=tk.X, padx=10)
        ttk.Button(riga_info, text="Cambia fattore…",
                   command=lambda: self._esegui_azione("f")).pack(side=tk.RIGHT)
        self.lbl_boll_info = ttk.Label(
            riga_info, text="Fattore: -   Specie rilevate: -",
            font=(_UI_FONT, 10), anchor=tk.W,
        )
        self.lbl_boll_info.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if not self.soglie:
            ttk.Label(info_frame,
                      text=f"ATTENZIONE: '{esportatori.SOGLIE_FILE_NOME}' non trovato, uso soglie di riserva.",
                      foreground=_COLORI["errore"], font=(_UI_FONT, 9), anchor=tk.W,
                      wraplength=340, justify=tk.LEFT).pack(fill=tk.X, padx=10)

        self._boll_scroll = _ScrollableFrame(tab)
        self._boll_scroll.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        # Niente larghezze fisse: la colonna della specie prende piu' spazio
        # (peso 3), i sette giorni e la media se lo dividono in parti uguali.
        griglia = self._boll_scroll.inner
        griglia.grid_columnconfigure(0, weight=3)
        for c in range(1, 9):
            griglia.grid_columnconfigure(c, weight=1, uniform="giorni")
        # Le celle si creano una volta sola, alla prima comparsa di una specie, e poi si
        # aggiornano sul posto: ricrearle a ogni inserimento costava ~9 widget per specie.
        for c, testo in enumerate(["Specie", *dominio.GIORNI_ABBREV, "Media"]):
            tk.Label(griglia, text=testo, font=(_UI_FONT, 9, "bold"),
                     relief=tk.RIDGE, bg="#e0e0e0").grid(row=0, column=c, sticky="nsew")
        self._boll_posizione = {riga[0]: i for i, riga in enumerate(dominio.BOLLETTINO_RIGHE)}
        self._boll_celle = {}       # {posizione: [9 Label]}
        self._boll_valori = {}      # {posizione: [(testo, bg, fg) delle 8 celle di dati]}
        self._boll_visibili = set()

    def _build_tab_codici(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Codici")

        ricerca_frame = ttk.Frame(tab, padding=(0, 6))
        ricerca_frame.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(ricerca_frame, text="Cerca:").pack(side=tk.LEFT, padx=(10, 4))
        self.entry_cerca_codici = ttk.Entry(ricerca_frame)
        self.entry_cerca_codici.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.entry_cerca_codici.bind("<KeyRelease>", self._filtra_tab_codici)

        azioni_frame = ttk.Frame(tab, padding=(0, 6))
        azioni_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(azioni_frame, text="Stampa elenco (Word)",
                   command=self._stampa_cheatsheet_codici).pack(side=tk.LEFT, padx=10)

        self.tree_codici = self._crea_tabella(
            tab, {"codice": "Cod.", "specie": "Specie", "voce": "A voce"})
        self._colonne_tabella(self.tree_codici, {"codice": "00000"}, ["specie", "voce"])
        self._parole_sconosciute = set()     # parole che il modello vocale non conosce
        vocabolario, _ = self._costruisci_vocabolario_vocale()
        self._forme_vocali = self._calcola_forme_vocali(vocabolario)

        self._popola_tab_codici()

    def _costruisci_vocabolario_vocale(self):
        """Legge la configurazione e ritorna (vocabolario, avvisi); aggiorna la
        parola di attivazione in uso."""
        extra = sessione.leggi_sinonimi_vocali(percorsi.CONFIG_FILE)
        self._voce_attivazione = sessione.leggi_parola_attivazione(percorsi.CONFIG_FILE)
        return dominio.costruisci_vocabolario_vocale(extra, self._voce_attivazione)

    def _calcola_forme_vocali(self, vocabolario):
        """{codice: 'ventiquattro, graminacee'}: cosa si puo' dire per ogni
        specie (il nome stesso della specie e' gia' nella colonna accanto).
        Dopo la prima attivazione della voce restano solo le forme che il
        modello vocale conosce davvero."""
        forme = {c: set() for c in dominio.TUTTI_CODICI}
        for frase, codice in vocabolario.items():
            if self._parole_sconosciute.intersection(frase.split()):
                continue
            # le trascrizioni ("be tu lace") si mostrano col nome vero ("betulacee")
            frase = dominio.TRASCRIZIONI_VOCALI.get(frase, frase)
            if frase != dominio.normalizza_parlato(dominio.CODICI_SPECIE[codice]):
                forme[codice].add(frase)
        return {c: ", ".join(sorted(f, key=lambda x: (len(x), x))) for c, f in forme.items()}

    def _popola_tab_codici(self, filtro=""):
        self.tree_codici.delete(*self.tree_codici.get_children())
        filtro = filtro.strip().lower()
        for codice in dominio.TUTTI_CODICI:
            specie = dominio.CODICI_SPECIE[codice]
            voce_forme = self._forme_vocali.get(codice, "")
            if filtro and not any(filtro in t.lower() for t in (codice, specie, voce_forme)):
                continue
            self.tree_codici.insert("", tk.END, values=(codice, specie, voce_forme))

    def _filtra_tab_codici(self, _event=None):
        self._popola_tab_codici(self.entry_cerca_codici.get())

    def _stampa_cheatsheet_codici(self):
        percorso = esportatori.genera_cheatsheet_codici(self.output_dir)
        if percorso is None:
            messagebox.showwarning(
                "python-docx non disponibile",
                "Per generare il foglio stampabile serve il modulo python-docx\n"
                "(su questa macchina: sudo apt install python3-docx, oppure\n"
                "sudo zypper install python3-python-docx su openSUSE).")
            return
        self._log(f"[OK] Elenco codici stampabile generato: {percorso}")
        messagebox.showinfo("Elenco generato", f"Foglio stampabile creato:\n{percorso}")

    # ================================================================
    # Lettura vocale (voce.py + dominio.interpreta_vocale)
    # ================================================================
    def _toggle_voce(self):
        try:
            self._toggle_voce_interno()
        finally:                      # l'interruttore mostra sempre lo stato reale
            self._var_voce.set(self._voce is not None)

    def _toggle_voce_interno(self):
        if self._voce is not None:
            self._voce_spegni()
            self._log("Voce disattivata.")
            return

        self._voce_vocabolario, avvisi = self._costruisci_vocabolario_vocale()
        grammatica = dominio.costruisci_grammatica(self._voce_vocabolario, self._voce_attivazione)
        attesa_ms = sessione.leggi_attesa_vocale(percorsi.CONFIG_FILE)
        vocabolario, attivazione = self._voce_vocabolario, self._voce_attivazione

        def frase_completa(parziale):
            # True se il parziale e' fatto solo di comandi validi (anche piu' di
            # seguito): la frase si puo' chiudere. "conta acero conta" e' ancora
            # incompleta: l'ultimo comando non e' finito.
            comandi = dominio.interpreta_vocale_sequenza(parziale, vocabolario, attivazione)
            return bool(comandi) and not any(
                isinstance(c, dominio.ComandoNonValido) for _, c in comandi)

        self._log("Carico il riconoscimento vocale...")
        self.root.update_idletasks()
        try:
            # Con la parola di attivazione la voce del PC non puo' essere scambiata
            # per un comando (le conferme non iniziano mai con "conta"): il microfono
            # resta aperto e non si perde cio' che l'operatore dice mentre il PC parla.
            asc, sintesi, sconosciute = voce.crea_voce(
                grammatica, frase_completa=frase_completa, attesa=attesa_ms / 1000,
                anti_eco=not self._voce_attivazione)
        except voce.VoceNonDisponibile as e:
            self._log("Lettura vocale non disponibile.")
            messagebox.showinfo("Lettura vocale non disponibile", str(e))
            return

        self._voce = (asc, sintesi)
        if sintesi.avviso:
            self._log(f"[voce] Attenzione: {sintesi.avviso}")
        for avviso in avvisi:
            self._log(f"[voce] Attenzione: {avviso}")
        if sconosciute:
            self._parole_sconosciute = set(sconosciute)
            self._forme_vocali = self._calcola_forme_vocali(vocabolario)
            self._popola_tab_codici(self.entry_cerca_codici.get())
            # le parole con una trascrizione (dominio.PRONUNCE_VOCALI) si capiscono lo stesso
            non_capite = [p for p in sconosciute if p not in dominio.PAROLE_CON_PRONUNCIA]
            if non_capite:
                self._log(f"[voce] {len(non_capite)} parole (soprattutto nomi latini) non sono nel "
                          "dizionario vocale e non verranno capite. Per quelle specie di' il numero "
                          "o il nome comune: vedi la colonna 'A voce' della scheda Codici.")
                self._log("[voce] Non capite: " + ", ".join(non_capite))
        inizio = f"{self._voce_attivazione} " if self._voce_attivazione else ""
        self._log(f"Voce ATTIVA. Di' ad esempio: \u00ab{inizio}ventiquattro\u00bb, "
                  f"\u00ab{inizio}graminacee per tre\u00bb, \u00ab{inizio}annulla\u00bb.")
        sintesi.parla("voce attiva")
        self._voce_after = self.root.after(50, self._poll_voce)

    def _toggle_talkback(self):
        self._talkback = not self._talkback
        self._var_talkback.set(self._talkback)
        if self._talkback:
            self._log("Ripeti a voce ATTIVO: il computer dice il nome di cio' che registra.")
        else:
            self._log("Ripeti a voce disattivato: le frasi vengono registrate in silenzio.")
        if self._voce is None:
            self._log("(La lettura vocale e' spenta: premi \"Voce\" per attivarla.)")

    def _voce_spegni(self):
        if self._voce_after is not None:
            self.root.after_cancel(self._voce_after)
            self._voce_after = None
        if self._voce is not None:
            asc, sintesi = self._voce
            self._voce = None
            sintesi.ferma()
            asc.ferma()
        self._var_voce.set(False)

    def _poll_voce(self):
        """Svuota la coda delle frasi capite (nel thread di Tk: tkinter non e'
        thread-safe, quindi il thread di ascolto si limita a riempire la coda)."""
        self._voce_after = None
        if self._voce is None:
            return
        asc, sintesi = self._voce
        if asc.errore:
            self._log(f"[voce] Errore del microfono: {asc.errore}")
            self._voce_spegni()
            return
        frase = asc.prossima_frase()
        while frase is not None:
            self._gestisci_frase_vocale(frase, sintesi)
            frase = asc.prossima_frase()
        self._voce_after = self.root.after(50, self._poll_voce)

    def _gestisci_frase_vocale(self, frase, sintesi):
        # Con un dialogo aperto (correggi, nota, conferme...) la voce non deve
        # toccare i dati: la frase viene scartata.
        if self.root.grab_current() is not None:
            return
        # Una frase puo' contenere piu' comandi di seguito ("conta acero conta
        # pioppo"): ogni "conta" chiude il precedente e ne apre uno nuovo.
        # Si eseguono nell'ordine detto; uno non valido non blocca gli altri.
        risposte = []
        for segmento, cmd in dominio.interpreta_vocale_sequenza(
                frase, self._voce_vocabolario, self._voce_attivazione):
            if isinstance(cmd, dominio.ComandoNonValido):
                if cmd.messaggio:
                    self._log(f"[voce] {cmd.messaggio}")
                    risposte.append("non capito")
                elif "[unk]" not in segmento:
                    # Aiuta a scoprire come il riconoscitore scrive i nomi che non
                    # conosce, per aggiungerli ai sinonimi.
                    self._log(f"[voce] Sentito ma non in elenco: \u00ab{segmento}\u00bb")
                continue    # rumore di fondo / fuori vocabolario: silenzio
            detto = self._esegui_comando(cmd, fonte="voce")
            # Il totale e' una domanda: si risponde sempre; il resto solo se richiesto.
            if detto and (self._talkback or isinstance(cmd, dominio.Totale)):
                risposte.append(detto)
        # Una sola frase parlata: la sintesi dice solo l'ultima in coda.
        if risposte and (self._talkback or any(r != "non capito" for r in risposte)):
            sintesi.parla(", ".join(risposte))

    # ================================================================
    # Selezione giorno
    # ================================================================
    def _pick_giorno(self, giorno_num):
        totale = self.settimana.totale_giorno(giorno_num)
        if totale > 0 and giorno_num != self.settimana.giorno_attivo:
            nome_g = dominio.GIORNI_NOMI[giorno_num].upper()
            if not messagebox.askyesno(
                    "Giorno gia' compilato",
                    f"{nome_g} contiene gia' {totale} osservazioni.\nContinuare aggiungendo dati?"):
                return
        self.settimana.attiva_giorno(giorno_num, journal=self.journal)
        self._imposta_giorno_ui(giorno_num)

    def _imposta_giorno_ui(self, giorno_num):
        for g, btn in self._giorno_bottoni.items():
            btn.config(style=("Giorno.Accent.TButton" if g == giorno_num else "Giorno.TButton"))
        self._aggiorna_etichetta_giorno(giorno_num)
        self._undo_consecutivi = 0
        self._imposta_entry_abilitata(True)
        self.entry.focus_set()

    def _aggiorna_etichetta_giorno(self, giorno_num):
        nome_giorno = dominio.GIORNI_NOMI[giorno_num].upper()
        data_str = self.settimana.data_str_giorno(giorno_num)
        totale = self.settimana.totale_giorno(giorno_num)
        self.lbl_giorno.config(text=f"{nome_giorno}  {data_str}  —  totale: {totale}")

    def _imposta_entry_abilitata(self, abilitata):
        self.entry.config(state=(tk.NORMAL if abilitata else tk.DISABLED))

    # ================================================================
    # Input: codici e comandi
    # ================================================================
    def _invia(self, _event=None):
        testo = self.entry.get()
        self.entry.delete(0, tk.END)
        if not testo.strip():
            return
        self._esegui_comando(dominio.interpreta_comando(testo))

    def _esegui_comando(self, cmd, fonte=""):
        """Esegue un comando gia' interpretato, da tastiera o da voce (stesso
        percorso: stesso modello, stesso journal). Ritorna il testo che la voce
        deve ripetere all'operatore, o None."""
        if self.settimana.giorno_attivo is None:
            self._log("Seleziona prima un giorno.")
            return None
        prefisso = f"[{fonte}] " if fonte else ""
        giorno_num = self.settimana.giorno_attivo

        if isinstance(cmd, dominio.Azione):
            if cmd.lettera == "u":
                # Da voce niente dialogo: bloccherebbe chi conta a mani libere e
                # le frasi dette nel frattempo verrebbero eseguite alla chiusura.
                return self._annulla(giorno_num, conferma=not fonte)
            self._esegui_azione(cmd.lettera)
            return None

        self._undo_consecutivi = 0

        if isinstance(cmd, dominio.Totale):
            totale = self.settimana.totale_giorno(giorno_num)
            self._log(f"{prefisso}Totale giornata: {totale}")
            return f"totale {totale}"

        if isinstance(cmd, dominio.Ripeti):
            if not self.settimana.ultimo_codice:
                self._log("Nessun codice precedente da ripetere.")
                return "niente da ripetere"
            codice, quantita = self.settimana.ultimo_codice, 1
        elif isinstance(cmd, dominio.Inserisci):
            codice, quantita = cmd.codice, cmd.quantita
        else:
            self._log(cmd.messaggio or "Comando non riconosciuto.")
            return "non capito" if fonte else None

        nuovo_val = self.settimana.inserisci(giorno_num, codice, quantita, journal=self.journal)
        self._modificato = True
        self._log(f"-> {prefisso}{dominio.descrivi_inserimento(codice, quantita)}  (totale giorno: {nuovo_val})")
        if self.settimana.beep:
            voce.beep()
        return self._testo_parlato_inserimento(codice, quantita)

    @staticmethod
    def _testo_parlato_inserimento(codice, quantita):
        """Cio' che la voce ripete: nome della specie (e quantita' se > 1)."""
        nome = dominio.normalizza_parlato(dominio.CODICI_SPECIE[codice])
        if quantita > 1:
            q = dominio.numero_in_parole(quantita) if quantita <= 59 else str(quantita)
            return f"{nome} per {q}"
        return nome

    def _esegui_azione(self, lettera):
        if self.settimana is None:
            return
        giorno_num = self.settimana.giorno_attivo

        if lettera == "h":
            self._mostra_aiuto()
        elif lettera == "r":
            self.notebook.select(1)
        elif lettera == "w":
            self.notebook.select(0)
        elif lettera == "l":
            self._mostra_storico()
        elif lettera == "b":
            self.settimana.beep = not self.settimana.beep
            self._var_beep.set(self.settimana.beep)
            self._log(f"Beep sonoro: {'ATTIVO' if self.settimana.beep else 'disattivo'}")
            if self.settimana.beep and not voce.beep():
                self._log("Non riesco a riprodurre il beep: manca un programma audio "
                          "(installa 'pulseaudio-utils' o 'alsa-utils').")
        elif lettera == "c":
            self._dialog_correggi()
        elif lettera == "n":
            self._dialog_nota()
        elif lettera == "f":
            self._dialog_fattore()
        elif lettera == "s":
            self._salva()
        elif lettera == "u":
            self._annulla(giorno_num)
        elif lettera == "d":
            self._chiudi_giornata()
        elif lettera == "q":
            self._esci_sessione()

    def _mostra_storico(self):
        if not self.settimana.storico:
            self._log("(nessun inserimento in questa sessione)")
            return
        self._log("Ultimi inserimenti:")
        for data, codice, _specie, quantita, ora in self.settimana.storico[-10:]:
            self._log(f"  {ora}  {dominio.descrivi_inserimento(codice, quantita)}  ({data})")

    def _annulla(self, giorno_num, conferma=True):
        if conferma:
            self._undo_consecutivi += 1
            if dominio.annullo_richiede_conferma(self._undo_consecutivi):
                if not messagebox.askyesno(
                        "Conferma",
                        f"Hai annullato {self._undo_consecutivi - 1} inserimenti di fila. Continuare?"):
                    self._undo_consecutivi = 0
                    return None
        ris = self.settimana.annulla(giorno_num, journal=self.journal)
        if ris is None:
            self._log("Nessun inserimento da annullare.")
            return "niente da annullare"
        codice, qty, _ = ris
        self._modificato = True
        self._log(f"<- Annullato: {dominio.descrivi_inserimento(codice, qty)}")
        return "annullato " + self._testo_parlato_inserimento(codice, qty)

    def _chiudi_giornata(self):
        if self.settimana.giorno_attivo is None:
            return
        nome_giorno = dominio.GIORNI_NOMI[self.settimana.giorno_attivo].upper()
        totale = self.settimana.totale_giorno(self.settimana.giorno_attivo)
        self._log(f"Chiusura {nome_giorno}: {totale} osservazioni.")
        self.settimana.giorno_attivo = None
        for btn in self._giorno_bottoni.values():
            btn.config(style="Giorno.TButton")
        self.lbl_giorno.config(text="Nessun giorno selezionato — sceglilo dalla barra sopra")
        self._imposta_entry_abilitata(False)

    # ================================================================
    # Dialoghi: correggi / nota
    # ================================================================
    def _prepara_dialogo(self, top, larghezza=None, altezza=None):
        """Dialogo centrato sulla finestra principale e modale: finche' e' aperto
        la voce e le scorciatoie non toccano i dati (vedi _gestisci_frase_vocale)."""
        top.transient(self.root)
        top.update_idletasks()
        larghezza = larghezza or top.winfo_reqwidth()
        altezza = altezza or top.winfo_reqheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - larghezza) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - altezza) // 3)
        top.geometry(f"{larghezza}x{altezza}+{x}+{y}")
        try:
            top.wait_visibility()       # senza, su Linux il grab fallisce
            top.grab_set()
        except tk.TclError:
            pass
        top.focus_set()

    def _dialog_correggi(self):
        if self.settimana is None:
            return
        top = tk.Toplevel(self.root)
        top.title("Correggi un giorno")
        top.minsize(380, 320)

        riga_giorno = ttk.Frame(top, padding=(0, 8))
        riga_giorno.pack(fill=tk.X, padx=12)
        ttk.Label(riga_giorno, text="Giorno:").pack(side=tk.LEFT)
        combo = ttk.Combobox(riga_giorno, state="readonly", width=14,
                             values=[f"{n}) {dominio.GIORNI_NOMI[n].upper()}" for n in range(1, 8)])
        combo.current((self.settimana.giorno_attivo or 1) - 1)
        combo.pack(side=tk.LEFT, padx=8)

        lista = tk.Listbox(top, font=(_UI_FONT, 10), relief=tk.FLAT, highlightthickness=1,
                           highlightbackground="#D0D0D0", activestyle="none",
                           selectbackground=_COLORI["accento"], selectforeground="#FFFFFF")
        lista.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        def _aggiorna_lista():
            lista.delete(0, tk.END)
            g = combo.current() + 1
            righe = self.settimana.righe_specie(g)
            if not righe:
                lista.insert(tk.END, "(nessun dato)")
            for codice, specie, val in righe:
                lista.insert(tk.END, f"[{codice}] {specie}: {val}")

        combo.bind("<<ComboboxSelected>>", lambda e: _aggiorna_lista())
        _aggiorna_lista()

        riga_valore = ttk.Frame(top, padding=(0, 8))
        riga_valore.pack(fill=tk.X, padx=12)
        ttk.Label(riga_valore, text="Nuovo valore:").pack(side=tk.LEFT)
        entry_val = ttk.Entry(riga_valore, width=8)
        entry_val.pack(side=tk.LEFT, padx=8)

        def _applica():
            sel = lista.curselection()
            if not sel:
                messagebox.showwarning("Nessuna selezione", "Seleziona una specie dall'elenco.", parent=top)
                return
            g = combo.current() + 1
            righe = self.settimana.righe_specie(g)
            if not righe or sel[0] >= len(righe):
                return
            codice, specie, vecchio = righe[sel[0]]
            try:
                nuovo_val = int(entry_val.get().strip())
                if nuovo_val < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Valore non valido", "Inserisci un numero intero >= 0.", parent=top)
                return
            self.settimana.correggi(g, codice, nuovo_val, journal=self.journal)
            self._modificato = True
            self._log(f"Corretto [{dominio.GIORNI_NOMI[g].upper()}]: [{codice}] {specie}: {vecchio} -> {nuovo_val}")
            entry_val.delete(0, tk.END)
            _aggiorna_lista()

        ttk.Button(riga_valore, text="Applica", style="Accent.TButton",
                   command=_applica).pack(side=tk.LEFT, padx=8)
        entry_val.bind("<Return>", lambda e: _applica())
        ttk.Button(top, text="Chiudi", command=top.destroy).pack(pady=(0, 12))
        top.bind("<Escape>", lambda e: top.destroy())
        self._prepara_dialogo(top, 420, 380)

    def _dialog_nota(self):
        if self.settimana is None or self.settimana.giorno_attivo is None:
            messagebox.showwarning("Nessun giorno attivo", "Seleziona prima un giorno.")
            return
        giorno_num = self.settimana.giorno_attivo
        testo = simpledialog.askstring(
            "Nota", f"Nota per {dominio.GIORNI_NOMI[giorno_num].upper()}:", parent=self.root)
        if not testo:
            return
        self.settimana.aggiungi_nota(giorno_num, testo, journal=self.journal)
        self._modificato = True
        self._log(f"Nota registrata: {testo}")

    def _dialog_fattore(self):
        if self.settimana is None:
            return
        attuale = self.settimana.fattore
        testo = simpledialog.askstring(
            "Fattore di conversione",
            "Fattore per trasformare la conta in concentrazione (granuli/m³).\n"
            "Vale per il bollettino e per il file salvato di questa settimana.\n\n"
            "Nuovo fattore:",
            initialvalue=f"{attuale:g}", parent=self.root)
        if testo is None or not testo.strip():
            return
        try:
            nuovo = dominio.interpreta_fattore(testo)
        except ValueError as e:
            messagebox.showerror("Valore non valido", str(e), parent=self.root)
            return
        if abs(nuovo - attuale) < 1e-12:
            return
        self.settimana.imposta_fattore(nuovo, journal=self.journal)
        self._modificato = True
        self._log(f"Fattore di conversione: {attuale:g} -> {nuovo:g}")

    # ================================================================
    # Salvataggio
    # ================================================================
    def _salva(self):
        """Esporta in .xlsx, chiedendo dove solo la prima volta. Ritorna il
        percorso, o None se l'utente annulla la scelta."""
        percorso = self.percorso_salvato
        if percorso is None:
            scelta = filedialog.asksaveasfilename(
                title="Salva file conta pollinica", initialdir=str(self.output_dir),
                initialfile=self.settimana.nome_file_default(), defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
            )
            if not scelta:
                self._log("Salvataggio annullato.")
                return None
            percorso = Path(scelta)
        try:
            esportatori.esporta_xlsx(self.settimana, percorso)
        except OSError as e:
            # Tipico su Windows: il file e' aperto in Excel e non si puo' riscrivere.
            self._log(f"ERRORE: non riesco a salvare {percorso}: {e}")
            messagebox.showerror(
                "Salvataggio non riuscito",
                f"Non riesco a salvare il file:\n{percorso}\n\n"
                "Se e' aperto in Excel o in un altro programma, chiudilo e riprova.\n"
                "I dati non sono persi: restano nella sessione.")
            return None
        self.percorso_salvato = percorso
        self._ora_salvataggio = datetime.now().strftime("%H:%M")
        self._modificato = False
        self._log(f"[OK] File salvato: {percorso}")
        return percorso

    # ================================================================
    # Uscita dalla sessione (comando 'q' / pulsante Esci)
    # ================================================================
    def _esci_sessione(self):
        risp = messagebox.askyesnocancel("Chiudere la settimana?",
                                         "Salvare il file prima di chiudere questa settimana?")
        if risp is None:
            return  # annulla, resta in sessione
        if risp:
            percorso = self._salva()
            if percorso is None:
                return  # ha annullato la scelta del percorso
            self.journal.elimina()
            self._menu_operazioni_aggiuntive(percorso.parent)
            self._torna_ad_avvio()
        else:
            conserva = messagebox.askyesno(
                "Uscita senza salvare",
                "I dati non ancora salvati in .xlsx restano nel file di sessione.\n"
                "Conservarlo per riprenderlo alla prossima apertura?",
            )
            if conserva:
                self.journal.chiudi()
                self._log(f"Il lavoro resta salvato in: {self.journal.path.name}")
            else:
                self.journal.elimina()
            self._torna_ad_avvio()

    def _menu_operazioni_aggiuntive(self, cartella):
        top = tk.Toplevel(self.root)
        top.title("Operazioni aggiuntive")
        ttk.Label(top, text="Vuoi generare anche:", padding=(18, 14, 18, 8),
                  style="Sezione.TLabel").pack(anchor="w")
        var_annuale = tk.BooleanVar(value=True)
        var_bollettini = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Aggiorna riepilogo annuale", variable=var_annuale).pack(anchor="w", padx=18, pady=2)
        ttk.Checkbutton(top, text="Genera bollettini Word (ITA/ENG)", variable=var_bollettini).pack(anchor="w", padx=18, pady=2)

        def _conferma():
            top.destroy()
            if var_annuale.get():
                percorso, n = esportatori.esporta_riepilogo_annuale(
                    self.settimana, cartella, scegli_duplicati=self._scegli_duplicati_dialogo,
                    avvisa=self._avvisa_fattore_diverso)
                if percorso:
                    self._log(f"[OK] Riepilogo annuale aggiornato: {percorso} ({n} giorni)")
                else:
                    self._log("Nessun dato da esportare nel riepilogo annuale.")
            if var_bollettini.get():
                soglie = esportatori.carica_soglie(cartella)
                if soglie is None:
                    self._log(f"ERRORE: file soglie '{esportatori.SOGLIE_FILE_NOME}' non trovato.")
                else:
                    creati = esportatori.genera_bollettini_word(self.settimana, cartella, soglie=soglie)
                    if creati:
                        for p in creati:
                            self._log(f"Bollettino: {p}")
                    else:
                        self._log("Nessun dato presente: bollettini Word non generati.")

        riga = ttk.Frame(top, padding=(0, 14))
        riga.pack(fill=tk.X)
        ttk.Button(riga, text="Conferma", style="Accent.TButton", command=_conferma).pack(side=tk.RIGHT, padx=18)
        ttk.Button(riga, text="Salta", command=top.destroy).pack(side=tk.RIGHT)
        top.bind("<Escape>", lambda e: top.destroy())
        self._prepara_dialogo(top, 380)
        top.wait_window(top)

    def _avvisa_fattore_diverso(self, messaggio):
        self._log("Attenzione: " + messaggio.replace("\n", " "))
        messagebox.showwarning("Fattore di conversione diverso", messaggio)

    def _scegli_duplicati_dialogo(self, data_str):
        top = tk.Toplevel(self.root)
        top.title("Giorno gia' presente")
        ttk.Label(top, text=f"Il giorno '{data_str}' \u00e8 gi\u00e0 presente nel riepilogo annuale.\n"
                            "Cosa vuoi fare?", padding=(18, 14, 18, 4), justify=tk.CENTER).pack()
        # Se la finestra viene chiusa per errore (tasto X) senza scegliere un
        # pulsante, il giorno non deve essere toccato: e' l'unica scelta sicura.
        scelta = {"valore": "annulla"}

        def _scegli(v):
            scelta["valore"] = v
            top.destroy()

        riga = ttk.Frame(top, padding=(14, 12))
        riga.pack()
        ttk.Button(riga, text="Sovrascrivi", command=lambda: _scegli("a")).pack(side=tk.LEFT, padx=6)
        ttk.Button(riga, text="Nuova riga", command=lambda: _scegli("b")).pack(side=tk.LEFT, padx=6)
        ttk.Button(riga, text="Somma", command=lambda: _scegli("c")).pack(side=tk.LEFT, padx=6)
        ttk.Button(riga, text="Annulla", style="Accent.TButton",
                   command=lambda: _scegli("annulla")).pack(side=tk.LEFT, padx=(24, 6))
        top.protocol("WM_DELETE_WINDOW", lambda: _scegli("annulla"))
        top.bind("<Escape>", lambda e: _scegli("annulla"))
        self._prepara_dialogo(top)
        top.wait_window(top)
        return scelta["valore"]

    def _torna_ad_avvio(self):
        self._voce_spegni()
        self._salva_layout()
        self.text_log = self.lbl_stato = self.lbl_file = None
        self.frame_lavoro.destroy()
        self.settimana = None
        self.journal = None
        self.percorso_salvato = None
        self._ora_salvataggio = None
        self._modificato = False
        self._giorno_bottoni = {}
        self._build_avvio()

    # ================================================================
    # Log testuale (sostituisce il vecchio terminale)
    # ================================================================
    # Colore della riga dedotto dall'inizio del testo (le chiamate non devono saperlo).
    _PREFISSI_LOG = (("->", "inserimento"), ("<-", "annullo"), ("[OK]", "ok"), ("ERRORE", "errore"),
                     ("Attenzione", "errore"), ("[voce]", "voce"))

    def _log(self, testo, tipo=None):
        if self.text_log is None:           # schermata iniziale: non c'e' il registro
            return
        if tipo is None:
            tipo = next((t for prefisso, t in self._PREFISSI_LOG if testo.startswith(prefisso)), None)
        self.text_log.config(state=tk.NORMAL)
        self.text_log.insert(tk.END, testo + "\n", tipo or ())
        n_righe = int(self.text_log.index("end-1c").split(".")[0])
        if n_righe > MAX_LINES_LOG:
            self.text_log.delete("1.0", f"{n_righe - MAX_LINES_LOG}.0")
        self.text_log.see(tk.END)
        self.text_log.config(state=tk.DISABLED)

    # ================================================================
    # Aggiornamento live delle tab (chiamato dal modello a ogni cambiamento:
    # nessun polling, nessuna rilettura di file, nessuna corsa fra thread)
    # ================================================================
    def _on_settimana_change(self, settimana):
        if settimana.giorno_attivo:
            self._aggiorna_etichetta_giorno(settimana.giorno_attivo)
        # Una frase vocale con piu' comandi notifica piu' volte di seguito: le
        # tab si ridisegnano una volta sola, appena Tk e' libero.
        self._tab_sporche.update(self._TAB_LIVE)
        if not self._refresh_in_attesa:
            self._refresh_in_attesa = True
            self.root.after_idle(self._refresh_tabs)

    @staticmethod
    def _nome_famiglia(codice):
        """Nome della famiglia senza note fra parentesi ('CORYLACEAE (somma c+o)'
        -> 'CORYLACEAE'): nelle tab lo spazio e' poco."""
        return re.sub(r"\s*\(.*\)", "", dominio.CODICI_SPECIE[codice])

    @staticmethod
    def _mostra_totale_famiglia(codice, conteggi):
        """True se per questa famiglia va mostrata la riga di totale: c'e' almeno
        un granulo di una specie della famiglia (se ci sono solo i 'dubbi' il
        totale coinciderebbe con la riga del codice e sarebbe un doppione)."""
        return codice in dominio.FAMIGLIE_SOMMA and any(
            sum(conteggi[c]) > 0 for c in dominio.FAMIGLIE_SOMMA[codice] if c != codice)

    @staticmethod
    def _celle_giorni(vals):
        return ["-" if v == 0 else str(v) for v in vals]

    # Indici delle tab live nel notebook (la tab Codici e' statica).
    _TAB_SETTIMANALE, _TAB_GIORNALIERO, _TAB_BOLLETTINO = 0, 1, 2
    _TAB_LIVE = (_TAB_SETTIMANALE, _TAB_GIORNALIERO, _TAB_BOLLETTINO)

    def _refresh_tabs(self):
        self._refresh_in_attesa = False
        self._aggiorna_tab_visibile()

    def _aggiorna_tab_visibile(self):
        """Ridisegna la tab aperta se i dati sono cambiati; le altre restano
        'sporche' e si aggiornano quando l'utente le apre."""
        if self.settimana is None:
            return
        try:
            corrente = self.notebook.index("current")
        except tk.TclError:             # schermata in chiusura
            return
        if corrente not in self._tab_sporche:
            return
        self._tab_sporche.discard(corrente)
        if corrente == self._TAB_SETTIMANALE:
            self._refresh_settimanale()
        elif corrente == self._TAB_GIORNALIERO:
            self._refresh_giornaliero()
        elif corrente == self._TAB_BOLLETTINO:
            self._refresh_bollettino()

    def _righe_specie(self, giornaliero):
        """[(iid, valori, tags)] per la tab Settimanale (totale) o Giornaliero
        (sette giorni): le specie con dati, con la riga di totale per famiglia."""
        conteggi = self.settimana.conteggi
        famiglie = dominio.totali_famiglia(conteggi)
        righe = []
        for codice in dominio.TUTTI_CODICI:
            vals = conteggi[codice]
            specie = dominio.CODICI_SPECIE[codice]
            if self._mostra_totale_famiglia(codice, conteggi):
                nome = self._nome_famiglia(codice)
                somma = famiglie[codice]
                celle = self._celle_giorni(somma) if giornaliero else (sum(somma),)
                righe.append(("Σ" + codice, ("Σ", f"{nome} (totale)", *celle), ("famiglia",)))
                specie = f"{nome} (dubbi)"
            if sum(vals) == 0:
                continue
            celle = self._celle_giorni(vals) if giornaliero else (sum(vals),)
            righe.append((codice, (codice, specie, *celle), ()))
        return righe

    def _aggiorna_tabella(self, tree, righe):
        """Allinea il Treeview a ``righe`` [(iid, valori, tags)]. Se l'elenco delle
        righe e' lo stesso (il caso normale: un granulo in piu' di una specie gia'
        presente) cambia solo il testo delle righe modificate; ricostruisce da
        capo solo quando una specie compare o sparisce."""
        cache = self._tree_cache.setdefault(tree, {})
        if list(tree.get_children()) != [iid for iid, _v, _t in righe]:
            tree.delete(*tree.get_children())
            cache.clear()
            for iid, valori, tags in righe:
                tree.insert("", tk.END, iid=iid, values=valori, tags=tags)
                cache[iid] = valori
            return
        for iid, valori, _tags in righe:
            if cache.get(iid) != valori:
                tree.item(iid, values=valori)
                cache[iid] = valori

    def _somme_pollini_spore(self):
        """([7 totali giornalieri dei pollini], [7 delle spore])."""
        pollini = set(dominio.POLLINI_CODICI)
        gp, gs = [0] * 7, [0] * 7
        for codice, vals in self.settimana.conteggi.items():
            somme = gp if codice in pollini else gs
            for i, v in enumerate(vals):
                somme[i] += v
        return gp, gs

    def _refresh_settimanale(self):
        self._aggiorna_tabella(self.tree_sett, self._righe_specie(giornaliero=False))
        gp, gs = self._somme_pollini_spore()
        self.lbl_s_pollini.config(text=f"Pollini: {sum(gp)}")
        self.lbl_s_spore.config(text=f"Spore: {sum(gs)}")
        self.lbl_s_totale.config(text=f"TOTALE: {sum(gp) + sum(gs)}")

    def _refresh_giornaliero(self):
        self._aggiorna_tabella(self.tree_giorn, self._righe_specie(giornaliero=True))
        gp, gs = self._somme_pollini_spore()

        def _fmt(vals):
            return "  ".join(f"{g}:{v}" for g, v in zip(dominio.GIORNI_ABBREV, vals) if v > 0)

        gt = [gp[i] + gs[i] for i in range(7)]
        self.lbl_g_pollini.config(text=f"Pollini: {_fmt(gp) or '-'}")
        self.lbl_g_spore.config(text=f"Spore:   {_fmt(gs) or '-'}")
        self.lbl_g_totale.config(text=f"TOTALE:  {_fmt(gt) or '-'}")

    def _refresh_bollettino(self):
        righe = dominio.righe_bollettino(self.settimana.conteggi, self.settimana.fattore, self.soglie)
        griglia = self._boll_scroll.inner
        presenti = set()
        for riga in righe:
            pos = self._boll_posizione[riga.nome_ita]
            presenti.add(pos)
            celle = self._boll_celle.get(pos)
            if celle is None:       # prima comparsa della specie: le 9 celle si creano ora
                celle = [tk.Label(griglia, text=riga.nome_ita, font=(_UI_FONT, 9),
                                  anchor="w", padx=3, relief=tk.RIDGE)]
                celle += [tk.Label(griglia, font=(_UI_FONT, 9, "bold" if c == 8 else "normal"),
                                   relief=tk.RIDGE) for c in range(1, 9)]
                self._boll_celle[pos] = celle
                self._boll_valori[pos] = [None] * 8
            # riga di griglia fissa = posizione in BOLLETTINO_RIGHE: l'ordine e' sempre quello del Word
            if pos not in self._boll_visibili:
                for c, w in enumerate(celle):
                    w.grid(row=pos + 1, column=c, sticky="nsew")
                self._boll_visibili.add(pos)

            nuovi = []
            for val, livello in zip(riga.concentrazioni, riga.livelli_giorno):
                bg, fg = dominio.LIVELLO_COLORE_GUI[livello]
                nuovi.append((f"{val:g}" if val else "-", bg, fg))
            bg, fg = dominio.LIVELLO_COLORE_GUI[riga.livello_media]
            nuovi.append((f"{riga.media:g}", bg, fg))
            for w, vecchio, nuovo in zip(celle[1:], self._boll_valori[pos], nuovi):
                if vecchio != nuovo:        # il testo/colore cambia solo nelle celle toccate
                    w.config(text=nuovo[0], bg=nuovo[1], fg=nuovo[2])
            self._boll_valori[pos] = nuovi

        for pos in self._boll_visibili - presenti:      # specie tornata a zero (annullo, correzione)
            for w in self._boll_celle[pos]:
                w.grid_remove()
        self._boll_visibili &= presenti

        self.lbl_boll_info.config(
            text=f"Fattore: {self.settimana.fattore:g}   Specie rilevate: {len(righe)}")

    # ================================================================
    # Chiusura applicazione
    # ================================================================
    def on_closing(self):
        if self.settimana is not None:
            if self._modificato or self.percorso_salvato is None:
                risp = messagebox.askyesnocancel(
                    "Uscire dal programma?",
                    "Ci sono dati non salvati in un file .xlsx definitivo "
                    "(restano comunque nel file di sessione).\nSalvare prima di uscire?")
                if risp is None:
                    return
                if risp:
                    if self._salva() is None:
                        return
                    self.journal.elimina()
                else:
                    self.journal.chiudi()
            else:
                self.journal.elimina()
        self._voce_spegni()
        self._salva_layout()
        self.root.destroy()


def main():
    if openpyxl is None:
        print("ERRORE: openpyxl non installato. Installa con:")
        print("  pip3 install openpyxl")
        sys.exit(1)

    root = tk.Tk()
    PollineCounterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
