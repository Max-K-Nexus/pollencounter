#!/usr/bin/env python3
"""
Conta Pollinica — interfaccia grafica.

Processo unico: la GUI chiama direttamente sessione.py/dominio.py/esportatori.py,
lo stesso dominio usato dalla CLI. Non lancia piu' polline_counter.py come
sottoprocesso: niente pty/pipe, niente marker __GUI_*__, niente rilettura
periodica di un autosave. Ogni inserimento aggiorna il modello in memoria e
notifica le tab con una chiamata diretta (nessun polling, nessuna corsa dati).
"""

import sys
import tkinter as tk
from datetime import timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import dominio
import esportatori
import percorsi
import sessione

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

MAX_LINES_LOG = 3000

HELP_TEXT = """Codici: 01-59 inserisce la specie corrispondente.
NNxQ inserisce Q occorrenze (es. 48x4 = 4 Alternaria).
.    ripete l'ultimo codice inserito.

Comandi (lettera + Invio nella casella di inserimento):
  r   riepilogo giornata corrente (apre la scheda Giornaliero)
  w   riepilogo settimanale (apre la scheda Settimanale)
  l   ultimi inserimenti
  c   correggi un giorno precedente
  n   aggiungi una nota per la giornata
  u   annulla ultimo inserimento
  b   attiva/disattiva il beep sonoro
  s   salva il file (senza uscire)
  d   chiudi la giornata corrente
  q   salva ed esci

Gli stessi comandi sono disponibili anche dai pulsanti sulla sinistra."""


class _ScrollableFrame(tk.Frame):
    """Frame con scrollbar verticale (per la scheda Bollettino, che disegna
    una griglia di celle colorate e non un semplice elenco)."""

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, bg="#ffffff")
        vscroll = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg="#ffffff")
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfig(self._window, width=e.width))
        self.canvas.configure(yscrollcommand=vscroll.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.bind("<Enter>", self._bind_wheel)
        self.canvas.bind("<Leave>", self._unbind_wheel)

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


class PollineCounterGUI:

    def __init__(self, root):
        self.root = root
        self.root.title("Conta Pollinica")
        self.root.geometry("1200x700")
        self.root.minsize(800, 400)

        self.output_dir = percorsi.EXE_DIR
        self.settimana = None
        self.journal = None
        self.nome_ripreso = None
        self.percorso_salvato = None
        self.soglie = {}
        self._modificato = False
        self._undo_consecutivi = 0
        self._giorno_bottoni = {}

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._carica_config_iniziale()
        self._build_avvio()

    # ================================================================
    # Configurazione cartella di lavoro
    # ================================================================
    def _carica_config_iniziale(self):
        import datetime as _dt
        anno = _dt.datetime.now().year
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
        import datetime as _dt
        sessione.salva_cartella_anno(percorsi.CONFIG_FILE, _dt.datetime.now().year, self.output_dir)
        self._aggiorna_lista_sessioni()

    # ================================================================
    # Schermata iniziale: scelta sessione/file/nuovo
    # ================================================================
    def _build_avvio(self):
        self.frame_avvio = tk.Frame(self.root, padx=16, pady=16)
        self.frame_avvio.pack(fill=tk.BOTH, expand=True)

        riga_cartella = tk.Frame(self.frame_avvio)
        riga_cartella.pack(fill=tk.X, pady=(0, 12))
        tk.Label(riga_cartella, text="Cartella di lavoro:", font=(_MONO_FONT, 10, "bold")).pack(side=tk.LEFT)
        self.lbl_cartella = tk.Label(riga_cartella, text=str(self.output_dir), font=(_MONO_FONT, 10))
        self.lbl_cartella.pack(side=tk.LEFT, padx=8)
        tk.Button(riga_cartella, text="Cambia cartella...", command=self._cambia_cartella).pack(side=tk.RIGHT)

        tk.Label(self.frame_avvio, text="Sessioni e file disponibili:",
                font=(_MONO_FONT, 10, "bold")).pack(anchor="w")

        cols = ("tipo", "nome", "info")
        self.tree_avvio = ttk.Treeview(self.frame_avvio, columns=cols, show="headings", height=12)
        self.tree_avvio.heading("tipo", text="Tipo")
        self.tree_avvio.heading("nome", text="Nome")
        self.tree_avvio.heading("info", text="Dettagli")
        self.tree_avvio.column("tipo", width=110, anchor=tk.W)
        self.tree_avvio.column("nome", width=320, anchor=tk.W)
        self.tree_avvio.column("info", width=260, anchor=tk.W)
        self.tree_avvio.pack(fill=tk.BOTH, expand=True, pady=8)
        self.tree_avvio.bind("<Double-1>", lambda e: self._avvio_riprendi_selezionata())

        riga_bottoni = tk.Frame(self.frame_avvio)
        riga_bottoni.pack(fill=tk.X, pady=(0, 12))
        tk.Button(riga_bottoni, text="Riprendi selezionata", command=self._avvio_riprendi_selezionata).pack(side=tk.LEFT)
        tk.Button(riga_bottoni, text="Importa file da un'altra cartella...", command=self._avvio_importa).pack(side=tk.LEFT, padx=8)

        riga_nuovo = tk.Frame(self.frame_avvio, pady=8)
        riga_nuovo.pack(fill=tk.X)
        tk.Label(riga_nuovo, text="Nuovo file — settimana:", font=(_MONO_FONT, 10, "bold")).pack(side=tk.LEFT)
        import datetime as _dt
        lun_corrente = dominio.lunedi_di(_dt.datetime.now())
        self.entry_settimana = tk.Entry(riga_nuovo, width=16, font=(_MONO_FONT, 10))
        self.entry_settimana.insert(0, lun_corrente.strftime("%d-%m-%Y"))
        self.entry_settimana.pack(side=tk.LEFT, padx=8)
        tk.Label(riga_nuovo, text="(es. 9-2-2026, 9/2/26, 9 feb 2026)",
                fg="#666666", font=(_MONO_FONT, 9)).pack(side=tk.LEFT)
        tk.Button(riga_nuovo, text="Nuovo file dal template", command=self._avvio_nuovo_file).pack(side=tk.RIGHT)

        self._aggiorna_lista_sessioni()

    def _aggiorna_lista_sessioni(self):
        self.lbl_cartella.config(text=str(self.output_dir))
        for item in self.tree_avvio.get_children():
            self.tree_avvio.delete(item)
        self._righe_avvio = []

        for path, info in sessione.recupera_sessioni(self.output_dir):
            self.tree_avvio.insert("", tk.END, values=(
                "Interrotta", f"settimana del {info['lunedi']}",
                f"{info['n_operazioni']} operazioni non salvate",
            ))
            self._righe_avvio.append(("journal", path))

        for f in sessione.cerca_file_ripresa(self.output_dir, esportatori.TEMPLATE_FILE.name):
            self.tree_avvio.insert("", tk.END, values=(
                "File salvato", f.name, sessione.conta_righe_log(f),
            ))
            self._righe_avvio.append(("xlsx", f))

    def _avvio_riprendi_selezionata(self):
        sel = self.tree_avvio.selection()
        if not sel:
            messagebox.showwarning("Nessuna selezione", "Seleziona una sessione o un file dall'elenco.")
            return
        idx = self.tree_avvio.index(sel[0])
        tipo, path = self._righe_avvio[idx]
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
        testo = self.entry_settimana.get().strip()
        dt = dominio.parse_data_flessibile(testo) if testo else None
        if testo and not dt:
            messagebox.showerror("Data non valida",
                                 "Non riconosco questa data. Prova con: 9-2-2026, 9/2/2026, 9 feb 2026.")
            return
        import datetime as _dt
        lunedi = dominio.lunedi_di(dt) if dt else dominio.lunedi_di(_dt.datetime.now())

        self.settimana = sessione.Settimana(lunedi)
        self.journal = sessione.Journal(self.output_dir / sessione.nome_journal(lunedi))
        self.journal.avvia(self.settimana)
        self.nome_ripreso = None
        self.percorso_salvato = None
        self._entra_in_lavoro()

    def _avvia_da_journal(self, path):
        try:
            self.settimana = sessione.ripristina_da_journal(path)
        except ValueError as e:
            messagebox.showerror("Sessione danneggiata", str(e))
            return
        self.journal = sessione.Journal.riprendi(path)
        self.nome_ripreso = self.settimana.nome_origine
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
                return
            self.settimana.lunedi = nuova

        self.nome_ripreso = self.settimana.nome_origine
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
        self._modificato = False
        self._build_lavoro()
        self.settimana.on_change(self._on_settimana_change)

        if self.settimana.giorno_attivo:
            self._imposta_giorno_ui(self.settimana.giorno_attivo)
        else:
            self._log("Seleziona un giorno per iniziare l'inserimento.")
            self._imposta_entry_abilitata(False)

        self._refresh_tabs()

    def _build_lavoro(self):
        self.pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=6, bg="#cccccc")
        self.pane.pack(fill=tk.BOTH, expand=True)

        # ── Pannello sinistro ──
        sinistra = tk.Frame(self.pane)
        self.pane.add(sinistra, stretch="always", width=650)

        barra_giorni = tk.Frame(sinistra, pady=4)
        barra_giorni.pack(side=tk.TOP, fill=tk.X)
        for g in range(1, 8):
            nome = dominio.GIORNI_ABBREV[g - 1]
            btn = tk.Button(barra_giorni, text=nome, width=5,
                            command=lambda g=g: self._pick_giorno(g))
            btn.pack(side=tk.LEFT, padx=2)
            self._giorno_bottoni[g] = btn

        self.lbl_giorno = tk.Label(sinistra, text="Nessun giorno selezionato",
                                   font=(_MONO_FONT, 11, "bold"), anchor="w")
        self.lbl_giorno.pack(side=tk.TOP, fill=tk.X, padx=6, pady=(2, 4))

        self.text_log = tk.Text(
            sinistra, wrap=tk.WORD, font=(_MONO_FONT, 11),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="#d4d4d4",
            state=tk.DISABLED, relief=tk.FLAT, padx=6, pady=6,
        )
        scroll_log = tk.Scrollbar(sinistra, command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=scroll_log.set)
        scroll_log.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_log.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        input_frame = tk.Frame(sinistra, bg="#2d2d2d")
        input_frame.pack(side=tk.BOTTOM, fill=tk.X)
        prompt_lbl = tk.Label(input_frame, text=" >> ", font=(_MONO_FONT, 11),
                              bg="#2d2d2d", fg="#00cc00")
        prompt_lbl.pack(side=tk.LEFT)
        self.entry = tk.Entry(input_frame, font=(_MONO_FONT, 11),
                              bg="#1e1e1e", fg="#d4d4d4",
                              insertbackground="#d4d4d4", relief=tk.FLAT)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4), pady=4)
        self.entry.bind("<Return>", self._invia)
        self.entry.focus_set()

        barra_azioni = tk.Frame(sinistra, pady=4)
        barra_azioni.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Button(barra_azioni, text="Aiuto (h)", command=lambda: self._esegui_azione("h")).pack(side=tk.LEFT, padx=2)
        tk.Button(barra_azioni, text="Correggi (c)", command=lambda: self._esegui_azione("c")).pack(side=tk.LEFT, padx=2)
        tk.Button(barra_azioni, text="Nota (n)", command=lambda: self._esegui_azione("n")).pack(side=tk.LEFT, padx=2)
        tk.Button(barra_azioni, text="Annulla (u)", command=lambda: self._esegui_azione("u")).pack(side=tk.LEFT, padx=2)
        self.btn_beep = tk.Button(barra_azioni, text="Beep: off", command=lambda: self._esegui_azione("b"))
        self.btn_beep.pack(side=tk.LEFT, padx=2)
        tk.Button(barra_azioni, text="Chiudi giornata (d)", command=lambda: self._esegui_azione("d")).pack(side=tk.RIGHT, padx=2)
        tk.Button(barra_azioni, text="Salva (s)", command=lambda: self._esegui_azione("s")).pack(side=tk.RIGHT, padx=2)
        tk.Button(barra_azioni, text="Esci (q)", command=lambda: self._esegui_azione("q")).pack(side=tk.RIGHT, padx=2)

        # ── Pannello destro ──
        destra = tk.Frame(self.pane)
        self.pane.add(destra, stretch="never", width=520)

        self.notebook = ttk.Notebook(destra)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("Summary.Treeview", font=(_MONO_FONT, 10), rowheight=22)
        style.configure("Summary.Treeview.Heading", font=(_MONO_FONT, 10, "bold"))
        if sys.platform == "darwin":
            style.map("Treeview", background=[], foreground=[])

        self._build_tab_settimanale()
        self._build_tab_giornaliero()
        self._build_tab_bollettino()
        self._build_tab_codici()

    def _build_tab_settimanale(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=" Settimanale ")

        columns = ("codice", "specie", "conteggio")
        self.tree_sett = ttk.Treeview(tab, columns=columns, show="headings", style="Summary.Treeview")
        self.tree_sett.heading("codice", text="Cod.")
        self.tree_sett.heading("specie", text="Specie")
        self.tree_sett.heading("conteggio", text="Tot.")
        self.tree_sett.column("codice", width=50, anchor=tk.CENTER)
        self.tree_sett.column("specie", width=220)
        self.tree_sett.column("conteggio", width=60, anchor=tk.CENTER)
        scroll = tk.Scrollbar(tab, command=self.tree_sett.yview)
        self.tree_sett.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_sett.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        totals = tk.Frame(tab, pady=8)
        totals.pack(side=tk.BOTTOM, fill=tk.X)
        self.lbl_s_pollini = tk.Label(totals, text="Pollini: 0", font=(_MONO_FONT, 11), anchor=tk.W)
        self.lbl_s_pollini.pack(fill=tk.X, padx=10)
        self.lbl_s_spore = tk.Label(totals, text="Spore: 0", font=(_MONO_FONT, 11), anchor=tk.W)
        self.lbl_s_spore.pack(fill=tk.X, padx=10)
        ttk.Separator(totals, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=4)
        self.lbl_s_totale = tk.Label(totals, text="TOTALE: 0", font=(_MONO_FONT, 12, "bold"), anchor=tk.W)
        self.lbl_s_totale.pack(fill=tk.X, padx=10)

    def _build_tab_giornaliero(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=" Giornaliero ")

        columns = ("codice", "specie", *[g.lower() for g in dominio.GIORNI_ABBREV])
        self.tree_giorn = ttk.Treeview(tab, columns=columns, show="headings", style="Summary.Treeview")
        self.tree_giorn.heading("codice", text="Cod.")
        self.tree_giorn.heading("specie", text="Specie")
        self.tree_giorn.column("codice", width=40, anchor=tk.CENTER)
        self.tree_giorn.column("specie", width=160)
        for g in dominio.GIORNI_ABBREV:
            col_id = g.lower()
            self.tree_giorn.heading(col_id, text=g)
            self.tree_giorn.column(col_id, width=40, anchor=tk.CENTER)
        scroll = tk.Scrollbar(tab, command=self.tree_giorn.yview)
        self.tree_giorn.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_giorn.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        totals = tk.Frame(tab, pady=8)
        totals.pack(side=tk.BOTTOM, fill=tk.X)
        self.lbl_g_pollini = tk.Label(totals, text="Pollini: -", font=(_MONO_FONT, 10), anchor=tk.W)
        self.lbl_g_pollini.pack(fill=tk.X, padx=10)
        self.lbl_g_spore = tk.Label(totals, text="Spore: -", font=(_MONO_FONT, 10), anchor=tk.W)
        self.lbl_g_spore.pack(fill=tk.X, padx=10)
        ttk.Separator(totals, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=4)
        self.lbl_g_totale = tk.Label(totals, text="TOTALE: -", font=(_MONO_FONT, 11, "bold"), anchor=tk.W)
        self.lbl_g_totale.pack(fill=tk.X, padx=10)

    def _build_tab_bollettino(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=" Bollettino ")

        self._boll_scroll = _ScrollableFrame(tab)
        self._boll_scroll.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        info_frame = tk.Frame(tab, pady=6)
        info_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.lbl_boll_info = tk.Label(
            info_frame, text="Fattore: -   Specie rilevate: -",
            font=(_MONO_FONT, 10), anchor=tk.W,
        )
        self.lbl_boll_info.pack(fill=tk.X, padx=10)
        if not self.soglie:
            tk.Label(info_frame,
                    text=f"ATTENZIONE: '{esportatori.SOGLIE_FILE_NOME}' non trovato, uso soglie di riserva.",
                    fg="#B00000", font=(_MONO_FONT, 9)).pack(fill=tk.X, padx=10)

    def _build_tab_codici(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=" Codici ")

        ricerca_frame = tk.Frame(tab, pady=6)
        ricerca_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(ricerca_frame, text="Cerca:").pack(side=tk.LEFT, padx=(10, 4))
        self.entry_cerca_codici = tk.Entry(ricerca_frame, font=(_MONO_FONT, 10))
        self.entry_cerca_codici.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        self.entry_cerca_codici.bind("<KeyRelease>", self._filtra_tab_codici)

        columns = ("codice", "specie")
        self.tree_codici = ttk.Treeview(tab, columns=columns, show="headings", style="Summary.Treeview")
        self.tree_codici.heading("codice", text="Cod.")
        self.tree_codici.heading("specie", text="Specie")
        self.tree_codici.column("codice", width=50, anchor=tk.CENTER)
        self.tree_codici.column("specie", width=220)
        scroll = tk.Scrollbar(tab, command=self.tree_codici.yview)
        self.tree_codici.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_codici.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self._popola_tab_codici()

        azioni_frame = tk.Frame(tab, pady=6)
        azioni_frame.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Button(azioni_frame, text="Stampa elenco (Word)",
                 command=self._stampa_cheatsheet_codici).pack(side=tk.LEFT, padx=10)

    def _popola_tab_codici(self, filtro=""):
        self.tree_codici.delete(*self.tree_codici.get_children())
        filtro = filtro.strip().lower()
        for codice in dominio.TUTTI_CODICI:
            specie = dominio.CODICI_SPECIE[codice]
            if filtro and filtro not in codice.lower() and filtro not in specie.lower():
                continue
            self.tree_codici.insert("", tk.END, values=(codice, specie))

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
            btn.config(relief=(tk.SUNKEN if g == giorno_num else tk.RAISED))
        nome_giorno = dominio.GIORNI_NOMI[giorno_num].upper()
        data_str = self.settimana.data_str_giorno(giorno_num)
        totale = self.settimana.totale_giorno(giorno_num)
        self.lbl_giorno.config(text=f"{nome_giorno}  {data_str}  —  totale: {totale}")
        self._undo_consecutivi = 0
        self._imposta_entry_abilitata(True)
        self.entry.focus_set()

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
        if self.settimana.giorno_attivo is None:
            self._log("Seleziona prima un giorno.")
            return

        cmd = dominio.interpreta_comando(testo)
        if isinstance(cmd, dominio.Azione):
            self._esegui_azione(cmd.lettera)
            return

        self._undo_consecutivi = 0
        giorno_num = self.settimana.giorno_attivo

        if isinstance(cmd, dominio.Ripeti):
            if not self.settimana.ultimo_codice:
                self._log("Nessun codice precedente da ripetere.")
                return
            codice, quantita = self.settimana.ultimo_codice, 1
        elif isinstance(cmd, dominio.Inserisci):
            codice, quantita = cmd.codice, cmd.quantita
        else:
            self._log(cmd.messaggio or "Comando non riconosciuto.")
            return

        specie = dominio.CODICI_SPECIE[codice]
        nuovo_val = self.settimana.inserisci(giorno_num, codice, quantita, journal=self.journal)
        self._modificato = True
        if quantita > 1:
            self._log(f"-> [{codice}] {specie} x{quantita}  (totale giorno: {nuovo_val})")
        else:
            self._log(f"-> [{codice}] {specie}  (totale giorno: {nuovo_val})")
        if self.settimana.beep:
            self.root.bell()

    def _esegui_azione(self, lettera):
        if self.settimana is None:
            return
        giorno_num = self.settimana.giorno_attivo

        if lettera == "h":
            messagebox.showinfo("Comandi", HELP_TEXT)
        elif lettera == "r":
            self.notebook.select(1)
        elif lettera == "w":
            self.notebook.select(0)
        elif lettera == "l":
            self._mostra_storico()
        elif lettera == "b":
            self.settimana.beep = not self.settimana.beep
            self.btn_beep.config(text=f"Beep: {'on' if self.settimana.beep else 'off'}")
            self._log(f"Beep sonoro: {'ATTIVO' if self.settimana.beep else 'disattivo'}")
            if self.settimana.beep:
                self.root.bell()
        elif lettera == "c":
            self._dialog_correggi()
        elif lettera == "n":
            self._dialog_nota()
        elif lettera == "s":
            self._salva_rapido()
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
        for data, codice, specie, quantita, ora in self.settimana.storico[-10:]:
            if quantita > 1:
                self._log(f"  {ora}  [{codice}] {specie} x{quantita}  ({data})")
            else:
                self._log(f"  {ora}  [{codice}] {specie}  ({data})")

    def _annulla(self, giorno_num):
        self._undo_consecutivi += 1
        if self._undo_consecutivi > 5 and (self._undo_consecutivi - 1) % 5 == 0:
            if not messagebox.askyesno(
                    "Conferma",
                    f"Hai annullato {self._undo_consecutivi - 1} inserimenti di fila. Continuare?"):
                self._undo_consecutivi = 0
                return
        ris = self.settimana.annulla(giorno_num, journal=self.journal)
        if ris is None:
            self._log("Nessun inserimento da annullare.")
            return
        codice, qty, _ = ris
        specie = dominio.CODICI_SPECIE[codice]
        self._modificato = True
        if qty > 1:
            self._log(f"<- Annullato: [{codice}] {specie} x{qty}")
        else:
            self._log(f"<- Annullato: [{codice}] {specie}")

    def _chiudi_giornata(self):
        if self.settimana.giorno_attivo is None:
            return
        nome_giorno = dominio.GIORNI_NOMI[self.settimana.giorno_attivo].upper()
        totale = self.settimana.totale_giorno(self.settimana.giorno_attivo)
        self._log(f"Chiusura {nome_giorno}: {totale} osservazioni.")
        self.settimana.giorno_attivo = None
        for btn in self._giorno_bottoni.values():
            btn.config(relief=tk.RAISED)
        self.lbl_giorno.config(text="Nessun giorno selezionato — sceglilo dalla barra sopra")
        self._imposta_entry_abilitata(False)

    # ================================================================
    # Dialoghi: correggi / nota
    # ================================================================
    def _dialog_correggi(self):
        if self.settimana is None:
            return
        top = tk.Toplevel(self.root)
        top.title("Correggi un giorno")
        top.transient(self.root)
        top.geometry("420x360")

        riga_giorno = tk.Frame(top, pady=8)
        riga_giorno.pack(fill=tk.X, padx=10)
        tk.Label(riga_giorno, text="Giorno:").pack(side=tk.LEFT)
        combo = ttk.Combobox(riga_giorno, state="readonly", width=14,
                             values=[f"{n}) {dominio.GIORNI_NOMI[n].upper()}" for n in range(1, 8)])
        combo.current((self.settimana.giorno_attivo or 1) - 1)
        combo.pack(side=tk.LEFT, padx=8)

        lista = tk.Listbox(top, font=(_MONO_FONT, 10))
        lista.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

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

        riga_valore = tk.Frame(top, pady=8)
        riga_valore.pack(fill=tk.X, padx=10)
        tk.Label(riga_valore, text="Nuovo valore:").pack(side=tk.LEFT)
        entry_val = tk.Entry(riga_valore, width=8)
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

        tk.Button(riga_valore, text="Applica", command=_applica).pack(side=tk.LEFT, padx=8)
        tk.Button(top, text="Chiudi", command=top.destroy).pack(pady=(0, 10))

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

    # ================================================================
    # Salvataggio
    # ================================================================
    def _nome_default(self):
        return self.nome_ripreso or f"Conta_Pollinica_{self.settimana.lunedi.strftime('%d-%m-%Y')}.xlsx"

    def _salva_rapido(self):
        percorso = self.percorso_salvato
        if percorso is None:
            percorso = filedialog.asksaveasfilename(
                title="Salva file conta pollinica", initialdir=str(self.output_dir),
                initialfile=self._nome_default(), defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
            )
            if not percorso:
                self._log("Salvataggio annullato.")
                return
            percorso = Path(percorso)
        esportatori.esporta_xlsx(self.settimana, percorso)
        self.percorso_salvato = percorso
        self._modificato = False
        self._log(f"[OK] File salvato: {percorso}")

    def _salva_come(self):
        percorso = filedialog.asksaveasfilename(
            title="Salva file conta pollinica", initialdir=str(self.output_dir),
            initialfile=self._nome_default(), defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
        )
        if not percorso:
            return None
        percorso = Path(percorso)
        esportatori.esporta_xlsx(self.settimana, percorso)
        self.percorso_salvato = percorso
        self._modificato = False
        return percorso

    # ================================================================
    # Uscita dalla sessione (comando 'q' / pulsante Esci)
    # ================================================================
    def _esci_sessione(self):
        risp = messagebox.askyesnocancel("Salvare ed uscire?", "Salvare il file prima di uscire dalla sessione?")
        if risp is None:
            return  # annulla, resta in sessione
        if risp:
            percorso = self.percorso_salvato or self._salva_come()
            if percorso is None:
                return  # ha annullato la scelta del percorso
            if self.percorso_salvato is None:
                esportatori.esporta_xlsx(self.settimana, percorso)
                self.percorso_salvato = percorso
            self._log(f"[OK] File salvato: {percorso}")
            self.journal.elimina()
            self._modificato = False
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
        top.transient(self.root)
        tk.Label(top, text="Vuoi generare anche:", padx=16, pady=10).pack(anchor="w")
        var_annuale = tk.BooleanVar(value=True)
        var_bollettini = tk.BooleanVar(value=True)
        tk.Checkbutton(top, text="Aggiorna riepilogo annuale", variable=var_annuale).pack(anchor="w", padx=16)
        tk.Checkbutton(top, text="Genera bollettini Word (ITA/ENG)", variable=var_bollettini).pack(anchor="w", padx=16)

        def _conferma():
            top.destroy()
            if var_annuale.get():
                percorso, n = esportatori.esporta_riepilogo_annuale(
                    self.settimana, cartella, scegli_duplicati=self._scegli_duplicati_dialogo)
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

        riga = tk.Frame(top, pady=10)
        riga.pack(fill=tk.X)
        tk.Button(riga, text="Conferma", command=_conferma).pack(side=tk.RIGHT, padx=16)
        tk.Button(riga, text="Salta", command=top.destroy).pack(side=tk.RIGHT)
        top.wait_window(top)

    def _scegli_duplicati_dialogo(self, data_str):
        top = tk.Toplevel(self.root)
        top.title("Giorno gia' presente")
        top.transient(self.root)
        tk.Label(top, text=f"Il giorno '{data_str}' e' gia' presente nel riepilogo annuale.",
                padx=16, pady=10).pack()
        # Se la finestra viene chiusa per errore (tasto X) senza scegliere un
        # pulsante, il giorno non deve essere toccato: e' l'unica scelta sicura.
        scelta = {"valore": "annulla"}

        def _scegli(v):
            scelta["valore"] = v
            top.destroy()

        riga = tk.Frame(top, pady=10)
        riga.pack()
        tk.Button(riga, text="Sovrascrivi", command=lambda: _scegli("a")).pack(side=tk.LEFT, padx=6)
        tk.Button(riga, text="Nuova riga", command=lambda: _scegli("b")).pack(side=tk.LEFT, padx=6)
        tk.Button(riga, text="Somma", command=lambda: _scegli("c")).pack(side=tk.LEFT, padx=6)
        tk.Button(riga, text="Annulla", command=lambda: _scegli("annulla")).pack(side=tk.LEFT, padx=(24, 6))
        top.protocol("WM_DELETE_WINDOW", lambda: _scegli("annulla"))
        top.wait_window(top)
        return scelta["valore"]

    def _torna_ad_avvio(self):
        self.pane.destroy()
        self.settimana = None
        self.journal = None
        self.nome_ripreso = None
        self.percorso_salvato = None
        self._modificato = False
        self._giorno_bottoni = {}
        self._build_avvio()

    # ================================================================
    # Log testuale (sostituisce il vecchio terminale)
    # ================================================================
    def _log(self, testo):
        self.text_log.config(state=tk.NORMAL)
        self.text_log.insert(tk.END, testo + "\n")
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
        self._refresh_tabs()
        if settimana.giorno_attivo:
            self._imposta_giorno_ui(settimana.giorno_attivo)

    def _refresh_tabs(self):
        s = self.settimana
        if s is None:
            return

        # Tab Settimanale
        self.tree_sett.delete(*self.tree_sett.get_children())
        tot_p, tot_s = 0, 0
        for codice, specie, val in s.righe_specie(None):
            self.tree_sett.insert("", tk.END, values=(codice, specie, val))
            if codice in dominio.POLLINI_CODICI:
                tot_p += val
            else:
                tot_s += val
        self.lbl_s_pollini.config(text=f"Pollini: {tot_p}")
        self.lbl_s_spore.config(text=f"Spore: {tot_s}")
        self.lbl_s_totale.config(text=f"TOTALE: {tot_p + tot_s}")

        # Tab Giornaliero
        self.tree_giorn.delete(*self.tree_giorn.get_children())
        gp, gs = [0] * 7, [0] * 7
        for codice in dominio.TUTTI_CODICI:
            vals = s.conteggi[codice]
            if not any(v > 0 for v in vals):
                continue
            display = [str(v) if v > 0 else "-" for v in vals]
            self.tree_giorn.insert("", tk.END, values=(codice, dominio.CODICI_SPECIE[codice], *display))
            for i in range(7):
                if codice in dominio.POLLINI_CODICI:
                    gp[i] += vals[i]
                else:
                    gs[i] += vals[i]

        def _fmt(vals):
            return "  ".join(f"{g}:{v}" for g, v in zip(dominio.GIORNI_ABBREV, vals) if v > 0)

        gt = [gp[i] + gs[i] for i in range(7)]
        self.lbl_g_pollini.config(text=f"Pollini: {_fmt(gp) or '-'}")
        self.lbl_g_spore.config(text=f"Spore:   {_fmt(gs) or '-'}")
        self.lbl_g_totale.config(text=f"TOTALE:  {_fmt(gt) or '-'}")

        # Tab Bollettino
        self._refresh_bollettino()

    def _refresh_bollettino(self):
        righe = dominio.righe_bollettino(self.settimana.conteggi, self.settimana.fattore, self.soglie)
        for child in self._boll_scroll.inner.winfo_children():
            child.destroy()

        larg_specie = 20
        header = ["Specie", *dominio.GIORNI_ABBREV, "Media"]
        for c, testo in enumerate(header):
            larghezza = larg_specie if c == 0 else 6
            tk.Label(self._boll_scroll.inner, text=testo, font=(_MONO_FONT, 9, "bold"),
                    width=larghezza, relief=tk.RIDGE, bg="#e0e0e0").grid(row=0, column=c, sticky="nsew")

        for r, riga in enumerate(righe, start=1):
            tk.Label(self._boll_scroll.inner, text=riga.nome_ita, font=(_MONO_FONT, 9),
                    width=larg_specie, anchor="w", relief=tk.RIDGE).grid(row=r, column=0, sticky="nsew")
            for c, (val, livello) in enumerate(zip(riga.concentrazioni, riga.livelli_giorno), start=1):
                bg, fg = dominio.LIVELLO_COLORE_GUI[livello]
                testo = f"{val:g}" if val else "-"
                tk.Label(self._boll_scroll.inner, text=testo, font=(_MONO_FONT, 9),
                        width=6, bg=bg, fg=fg, relief=tk.RIDGE).grid(row=r, column=c, sticky="nsew")
            bg, fg = dominio.LIVELLO_COLORE_GUI[riga.livello_media]
            tk.Label(self._boll_scroll.inner, text=f"{riga.media:g}", font=(_MONO_FONT, 9, "bold"),
                    width=6, bg=bg, fg=fg, relief=tk.RIDGE).grid(row=r, column=8, sticky="nsew")

        n_sp = len(righe)
        self.lbl_boll_info.config(text=f"Fattore: {self.settimana.fattore}   Specie rilevate: {n_sp}")

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
                    percorso = self.percorso_salvato or self._salva_come()
                    if percorso is None:
                        return
                    self._log(f"[OK] File salvato: {percorso}")
                    self.journal.elimina()
                else:
                    self.journal.chiudi()
            else:
                self.journal.elimina()
        self.root.destroy()


def main():
    if openpyxl is None:
        print("ERRORE: openpyxl non installato. Installa con:")
        print("  pip3 install openpyxl")
        sys.exit(1)

    root = tk.Tk()
    if sv_ttk is not None and sys.platform == "win32":
        sv_ttk.set_theme("light")
    PollineCounterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
