# 🌿 Pollencounter

Pollencounter is a pollen counting tool designed to support aerobiological monitoring and the production of official bulletins.

The main objective is to reduce manual labor on Excel files, standardize calculations, and simplify the management of readings and weekly reports. You do not need to be a programmer to use the Windows or macOS versions.

---

## ✨ Main Features

- **Automation** — Drastically reduces manual entry and repetitive calculations.
- **Standardization** — Uniform calculations to ensure the quality of aerobiological data.
- **Versatility** — Can be used via a Graphical User Interface (GUI) or Python scripts via Command Line (CLI).
- **Cross-platform** — Native support for Windows, macOS, and Linux.
- **Automatic Bulletins** — Generates pollen bulletins in Italian and English (`.docx`) with a concentration color scale.
- **Voice counting (optional)** — Record grains by voice without looking away from the microscope: say `conta ventiquattro` (code), `conta graminacee` (name) or `conta acero per tre` (quantity); `conta annulla` undoes the last entry. Works offline, repeats aloud what it understood, and ignores anything not preceded by the activation word *conta*. Italian only; see [Voice counting](#-voice-counting).
- **Crash recovery** — Every entry, undo, correction and note is written immediately to a session journal (`~sessione_<monday>.jsonl`); after a crash or power loss the interrupted session is offered for recovery at the next start.

---

## 📂 Repository Structure

```
pollencounter/
├── codice/                          # Source code and bundled files (single copy)
│   ├── polline_counter_gui.py       # GUI (tkinter) — recommended entry point
│   ├── polline_counter.py           # CLI (same commands, same data model)
│   ├── dominio.py                   # Pure domain logic: species codes, dates, thresholds, concentrations
│   ├── sessione.py                  # In-memory weekly model + crash-recovery journal
│   ├── esportatori.py               # Exports: weekly .xlsx, annual summary, Word bulletins
│   ├── percorsi.py                  # Path resolution (source vs. packaged executable)
│   ├── voce.py                      # Optional voice input/output (Vosk, sounddevice, pyttsx3)
│   ├── modelli/                     # Vosk Italian model for voice counting (not in git, see below)
│   ├── tests/                       # Automated tests (unittest)
│   ├── Polline_Template_Settimanale.xlsx        # Weekly Excel template
│   ├── concentrazioni_polliniche.xlsx           # Concentration thresholds (single source)
│   ├── ITA_Template_Bollettino_pubblicazione.docx  # Italian bulletin template
│   ├── ENG_Template_Bollettino_pubblicazione.docx  # English bulletin template
│   └── pollencounter.cfg.esempio    # Example configuration (work folder per year)
├── script_aiuto/                    # Linux launchers and maintenance utilities
│   ├── AVVIA_CONTA_POLLINICA.sh     # CLI launch
│   ├── AVVIA_CONTA_POLLINICA_GUI.sh # GUI launch
│   └── applica_formattazione.py     # Utility to update template formatting
├── mac/                             # macOS resources
│   ├── AVVIA_CONTA_POLLINICA.command     # CLI launch (double-click)
│   ├── AVVIA_CONTA_POLLINICA_GUI.command # GUI launch (double-click)
│   ├── build_app.sh                 # Script to build the .app bundle
│   └── ISTRUZIONI_MAC.txt           # macOS guide
├── windows/                         # Windows resources (no source copies here)
│   ├── AVVIA_CONTA_POLLINICA.bat    # Quick start with Python installed
│   ├── build_exe.bat                # Script to build the executable (dev)
│   └── ISTRUZIONI_WINDOWS.txt       # Windows guide
├── docs/                            # Technical review and design notes
├── riferimenti/                     # Historical tables and reference data
├── esempio di bollettino.pdf        # Example of the final bulletin
├── CHANGELOG.md                     # Change history
└── ISTRUZIONI.txt                   # General documentation (Italian)
```

---

## 🔄 Typical Workflow

1. **Start** — Open the application and choose the work folder for the season (remembered per year in `pollencounter.cfg`). Then resume an interrupted session, reopen a saved weekly file, or start a new week from the template.
2. **Counting** — Select the day and type one numeric code per pollen grain/spore observed (e.g. `24` = Gramineae, `48x4` = 4 Alternaria). Weekly, daily and bulletin views update live.
3. **Saving** — Save the week as `Conta_Pollinica_DD-MM-YYYY.xlsx` (command `s`, or `q` to save and exit).
4. **Output** — Optionally update the annual summary (`Riepilogo_Annuale_YYYY.xlsx`) and generate the Italian/English Word bulletins (see `esempio di bollettino.pdf`).

---

## 🚀 Getting Started

### 🪟 Windows Users (Non-technical)

The built executable does not require Python installation.

1. **The executable is not available for download**: it is not stored in the repository and it is not attached to the [Releases](https://github.com/Max-K-Nexus/pollencounter/releases) page. Build it yourself, once, on a Windows PC with Python installed: download the project (`Code → Download ZIP`), extract it and double-click `windows/build_exe.bat` (see `ISTRUZIONI_WINDOWS.txt`). The resulting `Conta_Pollinica.exe` can then be copied to any Windows PC and double-clicked.
   `build_exe.bat` includes [voice counting](#-voice-counting) **only if** the Vosk Italian model is in `codice/modelli/` when you run it; otherwise it builds the executable without voice.
2. Alternatively, with Python installed: download the project (`Code → Download ZIP`), extract it and double-click `windows/AVVIA_CONTA_POLLINICA.bat`.

### 🍎 macOS Users (Non-technical)

1. Download the project from GitHub (`Code → Download ZIP`) and extract it.
2. Open the `mac/` folder.
3. Read the `ISTRUZIONI_MAC.txt` file.
4. Launch the application by double-clicking `AVVIA_CONTA_POLLINICA_GUI.command` (the first time: right-click → Open).

Voice counting has been developed and tested on Linux; on macOS it is **untested** (the `.app` built by `mac/build_app.sh` includes it only if the model is present, see below).

### 🐍 Python Users (Developers)

Clone the repository:

```bash
git clone https://github.com/Max-K-Nexus/pollencounter.git
cd pollencounter
```

Install dependencies:

```bash
pip install openpyxl
# Optional — Word bulletins:
pip install python-docx
# Optional — Windows visual theme:
pip install sv-ttk
# Optional — voice counting (also needs the Vosk Italian model, see below):
pip install vosk sounddevice pyttsx3
```

On Debian/Ubuntu systems, tkinter may require separate installation:

```bash
sudo apt install python3-tk python3-docx
```

Run the script:

```bash
python3 codice/polline_counter_gui.py   # GUI (Recommended)
python3 codice/polline_counter.py       # CLI
```

Run the tests (from `codice/`):

```bash
cd codice && python3 -m unittest discover -s tests
```

---

## 🎙 Voice counting

An optional mode to record grains by voice, inspired by [EcoCount](https://doi.org/10.1177/2158244014537500) (Allen & Sewell, 2014). Press **Voce** in the GUI; every phrase must start with the activation word **conta**, otherwise it is ignored (so lab conversation is never recorded):

| Say | Effect |
|---|---|
| `conta ventiquattro` | records code 24 |
| `conta graminacee` | records a species by name |
| `conta acero per tre` | records 3 grains (2–20) |
| `conta annulla` / `conta ripeti` | undo / repeat the last entry |
| `conta totale` | reads the day's total aloud |

The computer repeats the name it understood, so a wrong recognition can be undone straight away. Saving, closing the day and quitting are keyboard/button only. *Prova voce* tells you what it understood without recording anything. Everything runs offline.

**Setup** (one-off): install `vosk sounddevice pyttsx3` (on Linux also `libportaudio2` and `espeak-ng`), then download `vosk-model-small-it-0.22` from <https://alphacephei.com/vosk/models> and unzip it into `codice/modelli/`. You need a microphone and, for the spoken confirmation, an Italian system voice. Many Latin names are not in the model's vocabulary: say the code or a common name instead (the **Codici** tab lists what each species accepts). Activation word and extra synonyms can be customised in `pollencounter.cfg` (see `codice/pollencounter.cfg.esempio` and `ISTRUZIONI.txt`).

**Status:** tested with a real microphone on Linux. The Windows executable was built and started under Wine only, and the macOS build is untested: please check the microphone and spoken confirmation on your system.

---

## ⚙️ Configuration and Conventions

**`.cfg` File** — `pollencounter.cfg` (created next to the script/executable on first start, see `codice/pollencounter.cfg.esempio`) stores the work folder chosen for each year. It is local to each machine and is not tracked by git.

**Concentration thresholds** — Bulletin levels (absent / low / medium / high) are read from `codice/concentrazioni_polliniche.xlsx`: edit that file to update them.

**File Naming Convention** — Keep the original names and folder structure. Weekly files use the `Conta_Pollinica_DD-MM-YYYY.xlsx` format (date of the Monday).

**Excel Templates** — Do not modify the structure of the `riepilogo_settimana` and `dati_grezzi` sheets. To update the visual formatting, use the dedicated script:

```bash
python3 script_aiuto/applica_formattazione.py
```

---

## 🛠 Troubleshooting (FAQ)

**The executable won't start?** Ensure you have correctly extracted the ZIP archive and that your antivirus is not blocking the executable file.

**Voice counting does not work?** Check that `vosk`, `sounddevice` and `pyttsx3` are installed and that the Italian model is in `codice/modelli/`; the **Voce** button explains what is missing. The program always works from the keyboard without it.

**Errors with Excel files?** Verify that you haven't modified or moved the column structure in the templates within the `codice/` folder.

**Using on Linux?** Use the scripts in the `script_aiuto/` folder.

**Using on macOS?** Use the scripts in the `mac/` folder. The GUI is compatible with macOS Tahoe (Tk 9.0) and earlier versions.

---

## 👥 Authors and Contributions

The Pollencounter project was developed by:

- **Simone Bettella** — Concept, original development, and calculation logic.
- **Massimiliano Iotti** — Maintenance, automation, documentation, and multi-platform support.

---

## 📜 License

The project is distributed under the **GNU General Public License v3.0 (GPL-3.0)**.

If you use Pollencounter in your project or report, please cite the authors:

> Pollencounter – developed by Simone Bettella and Massimiliano Iotti  
> [https://github.com/Max-K-Nexus/pollencounter](https://github.com/Max-K-Nexus/pollencounter)

---

## 📦 Requirements

| Dependency | Type | Notes |
|---|---|---|
| `openpyxl` | Mandatory | Excel reading/writing |
| `tkinter` | Mandatory (GUI) | On Debian: `sudo apt install python3-tk` |
| `python-docx` | Optional | Word bulletin generation |
| `sv-ttk` | Optional | Modern graphic theme (Windows only) |
| `vosk`, `sounddevice`, `pyttsx3` | Optional | Voice counting (also needs the Vosk Italian model; Linux: `libportaudio2`, `espeak-ng`) |
| `pyinstaller` | Development only | Windows/macOS executable build |
