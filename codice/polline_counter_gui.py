#!/usr/bin/env python3
"""
GUI wrapper per polline_counter.py — versione cross-platform.

Finestra tkinter con:
  - Terminale integrato (Text widget + Entry) che pilota lo script via pty (Linux)
    o subprocess stdin/stdout (Windows)
  - Tabella riepilogo live con tre schede: Settimanale, Giornaliero, Bollettino
"""

import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path

if sys.platform == "win32":
    import queue
else:
    import fcntl
    import pty
    import select
    import struct
    import termios

_MONO_FONT = "Courier New" if sys.platform == "win32" else "Monospace"

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = Path(sys.executable).parent
else:
    SCRIPT_DIR = Path(__file__).parent
SCRIPT_PATH = SCRIPT_DIR / "polline_counter.py"

# Importa costanti e helper da polline_counter (senza eseguire main)
sys.path.insert(0, str(SCRIPT_DIR))
from polline_counter import (
    CODICI_SPECIE, GIORNI_NOMI, SOGLIE_MAPPING,
    codice_to_row, giorno_to_col, leggi_valore, carica_soglie,
)

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import sv_ttk
except ImportError:
    sv_ttk = None

MAX_LINES = 5000
GIORNI_ABBREV = ["LUN", "MAR", "MER", "GIO", "VEN", "SAB", "DOM"]

# Colori (bg, fg) per i livelli di concentrazione nel tab Bollettino
_BOLL_COLORS = {
    "assente": ("#00B050", "#FFFFFF"),
    "bassa":   ("#FFD966", "#000000"),
    "media":   ("#F4B084", "#000000"),
    "alta":    ("#FF0000", "#FFFFFF"),
}


def _livello_conc(valore, soglia_tuple):
    """Ritorna il livello ('assente','bassa','media','alta') per un valore p/m³."""
    max_ass, max_bas, max_med = soglia_tuple
    if valore <= max_ass:
        return "assente"
    if valore <= max_bas:
        return "bassa"
    if valore <= max_med:
        return "media"
    return "alta"


class PollineCounterGUI:

    def __init__(self, root):
        self.root = root
        self.root.title("Conta Pollinica")
        self.root.geometry("1200x700")
        self.root.minsize(800, 400)

        self.master_fd = None
        self._output_queue = queue.Queue() if sys.platform == "win32" else None
        self.process = None
        self._tracked_file = None   # file Excel della sessione corrente
        self._sessione_attiva = False  # True solo dopo che file+giorno sono stati scelti
        self._refresh_running = False  # evita doppi timer concorrenti
        self._soglie = carica_soglie() or {}  # soglie per il bollettino
        self._marker_buf = ""  # buffer di accumulo per rilevamento marker

        self._build_ui()
        self._start_subprocess()
        self._poll_output()
        # Il refresh viene avviato solo quando file e giorno sono stati scelti

    # ── UI ────────────────────────────────────────────────────────

    def _build_ui(self):
        # PanedWindow orizzontale
        self.pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL,
                                   sashwidth=6, bg="#cccccc")
        self.pane.pack(fill=tk.BOTH, expand=True)

        # ── Pannello sinistro: terminale ──
        term_frame = tk.Frame(self.pane)
        self.pane.add(term_frame, stretch="always", width=650)

        self.text_output = tk.Text(
            term_frame, wrap=tk.WORD, font=(_MONO_FONT, 11),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="#d4d4d4",
            state=tk.DISABLED, relief=tk.FLAT, padx=6, pady=6,
        )
        scrollbar = tk.Scrollbar(term_frame, command=self.text_output.yview)
        self.text_output.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_output.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        input_frame = tk.Frame(term_frame, bg="#2d2d2d")
        input_frame.pack(side=tk.BOTTOM, fill=tk.X)

        prompt_lbl = tk.Label(input_frame, text=" >> ", font=(_MONO_FONT, 11),
                              bg="#2d2d2d", fg="#00cc00")
        prompt_lbl.pack(side=tk.LEFT)

        self.entry = tk.Entry(input_frame, font=(_MONO_FONT, 11),
                              bg="#1e1e1e", fg="#d4d4d4",
                              insertbackground="#d4d4d4", relief=tk.FLAT)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4), pady=4)
        self.entry.bind("<Return>", self._send_input)
        self.entry.focus_set()

        # ── Pannello destro: riepilogo con tab ──
        right_frame = tk.Frame(self.pane)
        self.pane.add(right_frame, stretch="never", width=520)

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Stile condiviso Treeview
        style = ttk.Style()
        style.configure("Summary.Treeview", font=(_MONO_FONT, 10), rowheight=22)
        style.configure("Summary.Treeview.Heading", font=(_MONO_FONT, 10, "bold"))

        # ── Tab 1: Settimanale ──
        self._build_tab_settimanale()

        # ── Tab 2: Giornaliero ──
        self._build_tab_giornaliero()

        # ── Tab 3: Bollettino ──
        self._build_tab_bollettino()

    def _build_tab_settimanale(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=" Settimanale ")

        columns = ("codice", "specie", "conteggio")
        self.tree_sett = ttk.Treeview(tab, columns=columns,
                                      show="headings", style="Summary.Treeview")
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

        self.lbl_s_pollini = tk.Label(totals, text="Pollini: 0",
                                      font=(_MONO_FONT, 11), anchor=tk.W)
        self.lbl_s_pollini.pack(fill=tk.X, padx=10)
        self.lbl_s_spore = tk.Label(totals, text="Spore: 0",
                                    font=(_MONO_FONT, 11), anchor=tk.W)
        self.lbl_s_spore.pack(fill=tk.X, padx=10)
        ttk.Separator(totals, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=4)
        self.lbl_s_totale = tk.Label(totals, text="TOTALE: 0",
                                     font=(_MONO_FONT, 12, "bold"), anchor=tk.W)
        self.lbl_s_totale.pack(fill=tk.X, padx=10)

    def _build_tab_giornaliero(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=" Giornaliero ")

        columns = ("codice", "specie",
                    "lun", "mar", "mer", "gio", "ven", "sab", "dom")
        self.tree_giorn = ttk.Treeview(tab, columns=columns,
                                       show="headings", style="Summary.Treeview")
        self.tree_giorn.heading("codice", text="Cod.")
        self.tree_giorn.heading("specie", text="Specie")
        self.tree_giorn.column("codice", width=40, anchor=tk.CENTER)
        self.tree_giorn.column("specie", width=160)
        for g in GIORNI_ABBREV:
            col_id = g.lower()
            self.tree_giorn.heading(col_id, text=g)
            self.tree_giorn.column(col_id, width=40, anchor=tk.CENTER)

        scroll = tk.Scrollbar(tab, command=self.tree_giorn.yview)
        self.tree_giorn.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_giorn.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        totals = tk.Frame(tab, pady=8)
        totals.pack(side=tk.BOTTOM, fill=tk.X)

        self.lbl_g_pollini = tk.Label(totals, text="Pollini: -",
                                      font=(_MONO_FONT, 10), anchor=tk.W)
        self.lbl_g_pollini.pack(fill=tk.X, padx=10)
        self.lbl_g_spore = tk.Label(totals, text="Spore: -",
                                    font=(_MONO_FONT, 10), anchor=tk.W)
        self.lbl_g_spore.pack(fill=tk.X, padx=10)
        ttk.Separator(totals, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=4)
        self.lbl_g_totale = tk.Label(totals, text="TOTALE: -",
                                     font=(_MONO_FONT, 11, "bold"), anchor=tk.W)
        self.lbl_g_totale.pack(fill=tk.X, padx=10)

    def _build_tab_bollettino(self):
        tab = tk.Frame(self.notebook)
        self.notebook.add(tab, text=" Bollettino ")

        cols = ("specie", "lun", "mar", "mer", "gio", "ven", "sab", "dom", "media")
        self.tree_boll = ttk.Treeview(tab, columns=cols,
                                      show="headings", style="Summary.Treeview")
        self.tree_boll.heading("specie", text="Specie")
        self.tree_boll.column("specie", width=160)
        for col_id, label in zip(cols[1:8], GIORNI_ABBREV):
            self.tree_boll.heading(col_id, text=label)
            self.tree_boll.column(col_id, width=42, anchor=tk.CENTER)
        self.tree_boll.heading("media", text="Media")
        self.tree_boll.column("media", width=52, anchor=tk.CENTER)

        # Tag colori per livello di concentrazione
        for nome, (bg, fg) in _BOLL_COLORS.items():
            self.tree_boll.tag_configure(nome, background=bg, foreground=fg)

        scroll = tk.Scrollbar(tab, command=self.tree_boll.yview)
        self.tree_boll.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_boll.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        info_frame = tk.Frame(tab, pady=6)
        info_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.lbl_boll_info = tk.Label(
            info_frame, text="Fattore: -   Specie rilevate: -",
            font=(_MONO_FONT, 10), anchor=tk.W,
        )
        self.lbl_boll_info.pack(fill=tk.X, padx=10)

    # ── Subprocess ────────────────────────────────────────────────

    def _start_subprocess(self):
        if sys.platform == "win32":
            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1"
            env["PYTHONIOENCODING"] = "utf-8"
            flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
            if getattr(sys, 'frozen', False):
                cmd = [sys.executable, "--cli"]
            else:
                cmd = [sys.executable, "-u", str(SCRIPT_PATH), "--gui"]
            self.process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                cwd=str(SCRIPT_DIR),
                bufsize=0,
                env=env,
                creationflags=flags,
            )
            self._reader_thread = threading.Thread(
                target=self._read_output_thread, daemon=True
            )
            self._reader_thread.start()
        else:
            master_fd, slave_fd = pty.openpty()
            winsize = struct.pack("HHHH", 40, 100, 0, 0)
            fcntl.ioctl(master_fd, termios.TIOCSWINSZ, winsize)
            self.process = subprocess.Popen(
                [sys.executable, str(SCRIPT_PATH), "--gui"],
                stdin=slave_fd, stdout=slave_fd, stderr=slave_fd,
                close_fds=True, cwd=str(SCRIPT_DIR),
            )
            os.close(slave_fd)
            self.master_fd = master_fd
            flags = fcntl.fcntl(master_fd, fcntl.F_GETFL)
            fcntl.fcntl(master_fd, fcntl.F_SETFL, flags | os.O_NONBLOCK)

    def _read_output_thread(self):
        """Solo Windows: legge l'output del processo in un thread separato."""
        try:
            while True:
                data = self.process.stdout.read(1)
                if not data:
                    break
                self._output_queue.put(data)
        except Exception:
            pass
        finally:
            self._output_queue.put(None)  # sentinel: processo terminato

    # ── I/O ───────────────────────────────────────────────────────

    def _poll_output(self):
        if sys.platform == "win32":
            self._poll_output_win32()
        else:
            self._poll_output_unix()

    _GUI_MARKERS = ("__GUI_ASKDIR__", "__GUI_ASKOPENFILE__", "__GUI_ASKSAVEFILE__")
    _MAX_MARKER_LEN = max(len(m) for m in _GUI_MARKERS)

    def _elabora_output(self, text, process_ended=False):
        """Processa testo raw: bell, ANSI, marker GUI, inserimento nel widget."""
        if "\a" in text:
            self.root.bell()
            text = text.replace("\a", "")
        text = re.sub(r"\x1b\[[0-9;]*[a-zA-Z]", "", text)
        display = self._handle_gui_markers(text, flush=process_ended)
        if display:
            self.text_output.config(state=tk.NORMAL)
            self.text_output.insert(tk.END, display)
            self._trim_output()
            self.text_output.see(tk.END)
            self.text_output.config(state=tk.DISABLED)
            self._detect_tracked_file(display)

    def _poll_output_win32(self):
        chunks = []
        process_ended = False
        try:
            while True:
                item = self._output_queue.get_nowait()
                if item is None:
                    process_ended = True
                    break
                chunks.append(item)
        except queue.Empty:
            pass
        if chunks:
            self._elabora_output(
                b"".join(chunks).decode("utf-8", errors="replace"),
                process_ended,
            )
        if process_ended:
            remaining = self._flush_marker_buf()
            if remaining:
                self.text_output.config(state=tk.NORMAL)
                self.text_output.insert(tk.END, remaining)
                self.text_output.config(state=tk.DISABLED)
            self._on_process_exit()
        else:
            self.root.after(50, self._poll_output)

    def _poll_output_unix(self):
        if self.master_fd is None:
            return
        process_alive = self.process and self.process.poll() is None
        try:
            ready, _, _ = select.select([self.master_fd], [], [], 0)
            if ready:
                data = os.read(self.master_fd, 8192)
                if data:
                    self._elabora_output(
                        data.decode("utf-8", errors="replace"),
                        not process_alive,
                    )
        except OSError:
            pass

        if process_alive:
            self.root.after(50, self._poll_output)
        else:
            remaining = self._flush_marker_buf()
            if remaining:
                self.text_output.config(state=tk.NORMAL)
                self.text_output.insert(tk.END, remaining)
                self.text_output.config(state=tk.DISABLED)
            self._on_process_exit()

    def _send_to_stdin(self, text):
        """Invia una riga al processo (cross-platform)."""
        if sys.platform == "win32":
            if self.process and self.process.stdin:
                try:
                    self.process.stdin.write((text + "\n").encode("utf-8"))
                    self.process.stdin.flush()
                except OSError:
                    pass
        else:
            if self.master_fd is not None:
                try:
                    os.write(self.master_fd, (text + "\n").encode("utf-8"))
                except OSError:
                    pass

    def _send_input(self, _event=None):
        text = self.entry.get()
        self.entry.delete(0, tk.END)
        self._send_to_stdin(text)

    def _handle_gui_markers(self, new_text, flush=False):
        """Accumula testo, intercetta marker completi, ritorna testo sicuro da mostrare.

        Il buffer interno trattiene i frammenti finali che potrebbero essere
        l'inizio di un marker spezzato tra due poll. Quando un marker completo
        viene trovato, lo rimuove dal buffer e apre il dialogo corrispondente.
        """
        self._marker_buf += new_text
        display_parts = []

        # Cerca e processa tutti i marker completi nel buffer
        while True:
            earliest_pos = -1
            earliest_marker = None
            for marker in self._GUI_MARKERS:
                pos = self._marker_buf.find(marker)
                if pos != -1 and (earliest_pos == -1 or pos < earliest_pos):
                    earliest_pos = pos
                    earliest_marker = marker
            if earliest_marker is None:
                break
            # Testo prima del marker → va mostrato
            display_parts.append(self._marker_buf[:earliest_pos])
            self._marker_buf = self._marker_buf[earliest_pos + len(earliest_marker):]
            # Apri il dialogo appropriato
            self.root.lift()
            self.root.focus_force()
            if earliest_marker == "__GUI_ASKDIR__":
                path = filedialog.askdirectory(
                    parent=self.root,
                    title="Scegli cartella di salvataggio",
                    initialdir=str(SCRIPT_DIR),
                )
            elif earliest_marker == "__GUI_ASKSAVEFILE__":
                path = filedialog.asksaveasfilename(
                    parent=self.root,
                    title="Salva file conta pollinica",
                    initialdir=str(SCRIPT_DIR),
                    defaultextension=".xlsx",
                    filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
                )
            else:
                path = filedialog.askopenfilename(
                    parent=self.root,
                    title="Importa file conta pollinica",
                    initialdir=str(SCRIPT_DIR),
                    filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")],
                )
            self._send_to_stdin(path if path else "")

        if flush:
            # Processo terminato: svuota tutto il buffer
            display_parts.append(self._marker_buf)
            self._marker_buf = ""
        else:
            # Trattieni la coda che potrebbe essere un marker parziale.
            # Tutto il testo prima dell'ultimo possibile inizio di marker
            # e' sicuro da mostrare.
            safe_end = len(self._marker_buf)
            for length in range(min(self._MAX_MARKER_LEN - 1, len(self._marker_buf)), 0, -1):
                tail = self._marker_buf[-length:]
                if any(m.startswith(tail) for m in self._GUI_MARKERS):
                    safe_end = len(self._marker_buf) - length
                    break
            display_parts.append(self._marker_buf[:safe_end])
            self._marker_buf = self._marker_buf[safe_end:]

        return "".join(display_parts)

    def _flush_marker_buf(self):
        """Svuota il buffer marker e ritorna il testo residuo."""
        remaining = self._marker_buf
        self._marker_buf = ""
        return remaining

    def _trim_output(self):
        line_count = int(self.text_output.index("end-1c").split(".")[0])
        if line_count > MAX_LINES:
            self.text_output.delete("1.0", f"{line_count - MAX_LINES}.0")

    # ── Tracking del file corrente ───────────────────────────────

    def _detect_tracked_file(self, text):
        """Analizza l'output dello script per capire quale file sta usando."""
        aggiorna = False

        # Rileva "Ripreso: /percorso/completo/nomefile.xlsx"
        match = re.search(r"Ripreso:\s*(.+\.xlsx)", text)
        if match:
            path = Path(match.group(1).strip())
            if path.exists():
                self._tracked_file = path
                aggiorna = True

        # Rileva "[auto-salvato]: /percorso/completo/~autosave_*.xlsx"
        match = re.search(r"\[auto-salvato\]:\s*(.+\.xlsx)", text)
        if match:
            path = Path(match.group(1).strip())
            if path.exists():
                self._tracked_file = path
            aggiorna = True

        # Rileva "File salvato: /percorso/completo/nomefile.xlsx" (salvataggio definitivo)
        match = re.search(r"File salvato:\s*(.+\.xlsx)", text)
        if match:
            path = Path(match.group(1).strip())
            if path.exists():
                self._tracked_file = path
                aggiorna = True

        # Rileva "Sessione sospesa. File salvato: /percorso/completo/nomefile.xlsx"
        match = re.search(r"Sessione sospesa.*?:\s*(.+\.xlsx)", text)
        if match:
            path = Path(match.group(1).strip())
            if path.exists():
                self._tracked_file = path
                aggiorna = True

        # Rileva inizio sessione giorno (giorno e file sono stati scelti)
        if re.search(r"Giorno:\s+\w+", text):
            self._sessione_attiva = True
            aggiorna = True

        # Rileva fine sessione
        if "Sessione terminata" in text or "Sessione sospesa" in text:
            self._sessione_attiva = False

        # Avvia il refresh solo se sessione attiva e non già in esecuzione
        if aggiorna and self._sessione_attiva and not self._refresh_running:
            self._refresh_running = True
            self.root.after(500, self._refresh_summary)

    # ── Riepilogo live ────────────────────────────────────────────

    def _refresh_summary(self):
        # NON resettare _refresh_running qui: rimane True finche' il thread
        # e' in esecuzione, cosi' _detect_tracked_file non schedula duplicati.

        # Refresh solo se la sessione è attiva (file + giorno scelti)
        if not self._sessione_attiva:
            self._refresh_running = False
            return

        if openpyxl is None:
            self._refresh_running = False
            return

        if self._tracked_file is None or not self._tracked_file.exists():
            # File non ancora noto o non ancora scritto dal thread autosave: riprova
            self.root.after(3000, self._refresh_summary)
            return

        # Lettura Excel in thread separato per non bloccare il main thread
        filepath = self._tracked_file
        threading.Thread(target=self._leggi_dati_thread,
                         args=(filepath,), daemon=True).start()

    def _leggi_dati_thread(self, filepath):
        """Legge il file Excel in background. Aggiorna la UI nel main thread."""
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb["riepilogo_settimana"]
            # Ricarica soglie dal workbook (foglio "soglie" integrato)
            soglie_aggiornate = carica_soglie(wb)
            if soglie_aggiornate:
                self._soglie = soglie_aggiornate
            dati = self._raccogli_dati(ws)
            wb.close()
            # Torna al main thread per aggiornare i widget
            self.root.after(0, lambda: self._applica_dati(dati))
        except Exception:
            pass
        finally:
            # Riprogramma il prossimo ciclo nel main thread
            self.root.after(0, self._schedula_prossimo_refresh)

    def _schedula_prossimo_refresh(self):
        if self._sessione_attiva:
            # _refresh_running rimane True: il prossimo timer e' gia' in coda
            self.root.after(3000, self._refresh_summary)
        else:
            self._refresh_running = False

    def _raccogli_dati(self, ws):
        """Raccoglie tutti i dati dal foglio (eseguito nel thread background)."""
        righe_sett = []
        totale_pollini_s = 0
        totale_spore_s = 0

        righe_giorn = []
        tot_pollini_g = [0] * 7
        tot_spore_g = [0] * 7

        for codice_str, specie in CODICI_SPECIE.items():
            row = codice_to_row(codice_str)
            if row is None:
                continue
            vals = [leggi_valore(ws, row, giorno_to_col(g)) for g in range(1, 8)]
            total = sum(vals)
            n = int(codice_str)

            if total > 0:
                righe_sett.append((codice_str, specie, total))
                if n <= 47:
                    totale_pollini_s += total
                else:
                    totale_spore_s += total

            if any(v > 0 for v in vals):
                display = [str(v) if v > 0 else "-" for v in vals]
                righe_giorn.append((codice_str, specie, *display))
                for i in range(7):
                    if n <= 47:
                        tot_pollini_g[i] += vals[i]
                    else:
                        tot_spore_g[i] += vals[i]

        # ── Bollettino: concentrazioni p/m³ per famiglia ──
        fattore_val = ws["Q3"].value
        fattore = float(fattore_val) if isinstance(fattore_val, (int, float)) and fattore_val > 0 else 0.4

        righe_boll = []
        for codice, famiglia_soglia in SOGLIE_MAPPING.items():
            row_riep = codice_to_row(codice)
            if row_riep is None:
                continue
            vals_conta = [leggi_valore(ws, row_riep, giorno_to_col(g)) for g in range(1, 8)]
            if all(v == 0 for v in vals_conta):
                continue
            conc = [round(v * fattore, 1) for v in vals_conta]
            media = round(sum(conc) / 7.0, 1)
            nome = CODICI_SPECIE.get(codice, famiglia_soglia)
            soglia_tuple = self._soglie.get(famiglia_soglia, (0.9, 19.9, 39.9))
            livello = _livello_conc(media, soglia_tuple)
            display_conc = [str(v) if v > 0 else "-" for v in conc]
            righe_boll.append((nome, *display_conc, str(media), livello))

        return {
            "sett_righe": righe_sett,
            "sett_pollini": totale_pollini_s,
            "sett_spore": totale_spore_s,
            "giorn_righe": righe_giorn,
            "giorn_pollini": tot_pollini_g,
            "giorn_spore": tot_spore_g,
            "boll_righe": righe_boll,
            "boll_fattore": fattore,
        }

    def _applica_dati(self, dati):
        """Aggiorna i widget Treeview con i dati raccolti (main thread)."""
        # Tab settimanale
        ch = self.tree_sett.get_children()
        if ch:
            self.tree_sett.delete(*ch)
        for riga in dati["sett_righe"]:
            self.tree_sett.insert("", tk.END, values=riga)
        p = dati["sett_pollini"]
        s = dati["sett_spore"]
        self.lbl_s_pollini.config(text=f"Pollini: {p}")
        self.lbl_s_spore.config(text=f"Spore: {s}")
        self.lbl_s_totale.config(text=f"TOTALE: {p + s}")

        # Tab giornaliero
        ch = self.tree_giorn.get_children()
        if ch:
            self.tree_giorn.delete(*ch)
        for riga in dati["giorn_righe"]:
            self.tree_giorn.insert("", tk.END, values=riga)

        def _fmt(vals):
            return "  ".join(f"{g}:{v}" for g, v in zip(GIORNI_ABBREV, vals) if v > 0)

        gp = dati["giorn_pollini"]
        gs = dati["giorn_spore"]
        gt = [gp[i] + gs[i] for i in range(7)]
        self.lbl_g_pollini.config(text=f"Pollini: {_fmt(gp) or '-'}")
        self.lbl_g_spore.config(text=f"Spore:   {_fmt(gs) or '-'}")
        self.lbl_g_totale.config(text=f"TOTALE:  {_fmt(gt) or '-'}")

        # Tab bollettino
        ch = self.tree_boll.get_children()
        if ch:
            self.tree_boll.delete(*ch)
        for riga in dati["boll_righe"]:
            *vals, livello = riga
            self.tree_boll.insert("", tk.END, values=vals, tags=(livello,))
        n_sp = len(dati["boll_righe"])
        self.lbl_boll_info.config(
            text=f"Fattore: {dati['boll_fattore']}   Specie rilevate: {n_sp}"
        )

    # ── Chiusura ──────────────────────────────────────────────────

    def _on_process_exit(self):
        self.text_output.config(state=tk.NORMAL)
        self.text_output.insert(tk.END, "\n--- Processo terminato ---\n")
        self.text_output.config(state=tk.DISABLED)
        self.entry.config(state=tk.DISABLED)
        if sys.platform != "win32" and self.master_fd is not None:
            try:
                os.close(self.master_fd)
            except OSError:
                pass
            self.master_fd = None

    def on_closing(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
        if sys.platform != "win32" and self.master_fd is not None:
            try:
                os.close(self.master_fd)
            except OSError:
                pass
            self.master_fd = None
        self.root.destroy()


def main():
    root = tk.Tk()
    if sv_ttk is not None:
        sv_ttk.set_theme("light")
    app = PollineCounterGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    if "--cli" in sys.argv:
        from polline_counter import main as cli_main
        cli_main()
    else:
        main()
