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
- **Autosave** — Automatic saving every 5 entries; data is never lost.

---

## 📂 Repository Structure

```
pollencounter/
├── codice/                          # Main scripts and reference files
│   ├── polline_counter.py           # Processing logic (CLI, cross-platform)
│   ├── polline_counter_gui.py       # GUI version (tkinter)
│   ├── Polline_Template_Settimanale.xlsx        # Base template for calculations
│   ├── concentrazioni_polliniche.xlsx           # Bulletin thresholds (fallback)
│   ├── ITA_Template_Bollettino_pubblicazione.docx  # Italian bulletin template
│   ├── ENG_Template_Bollettino_pubblicazione.docx  # English bulletin template
│   └── pollencounter.cfg            # Folder configuration and parameters
├── script_aiuto/                    # Startup and maintenance utilities (Linux)
│   ├── AVVIA_CONTA_POLLINICA.sh     # CLI Launch
│   ├── AVVIA_CONTA_POLLINICA_GUI.sh # GUI Launch
│   ├── applica_formattazione.py     # Utility to update template formatting
│   └── setup_bollettino_template.py # Template configuration utility
├── mac/                             # macOS specific resources
│   ├── AVVIA_CONTA_POLLINICA.sh     # CLI Launch
│   ├── AVVIA_CONTA_POLLINICA_GUI.sh # GUI Launch
│   ├── build_app.sh                 # Script to compile the .app bundle
│   └── ISTRUZIONI_MAC.txt           # macOS specific guide
├── windows/                         # Windows specific resources
│   ├── Conta_Pollinica.exe          # Pre-compiled executable (ready to use)
│   ├── AVVIA_CONTA_POLLINICA.bat    # Quick start with Python installed
│   ├── build_exe.bat                # Script to compile the executable (dev)
│   └── ISTRUZIONI_WINDOWS.txt       # Windows specific guide
├── riferimenti/                     # Historical tables and reference data
├── esempio di bolletino.pdf         # Final output example
├── CHANGELOG.md                     # Change history
└── ISTRUZIONI.txt                   # General documentation
```

---

## 🔄 Typical Workflow

1. **Data Collection** — The operator fills out the weekly Excel files in the `letture_settimanali/` folder, following the `Conta_Pollinica_DD-MM-YYYY.xlsx` naming convention.
2. **Processing** — Run the application (exe, app, or Python script). The program reads the input files, templates, and configuration in the `codice/` folder.
3. **Output** — The tool generates updated Excel files and Word bulletins ready for publication (as shown in `esempio di bolletino.pdf`).

---

## 🚀 Getting Started

### 🪟 Windows Users (Non-technical)

This method does not require Python installation.

1. Download the project from GitHub (`Code → Download ZIP`) and extract it.
2. Open the `windows/` folder.
3. Read the `ISTRUZIONI_WINDOWS.txt` file.
4. Launch the application by double-clicking `Conta_Pollinica.exe` or `AVVIA_CONTA_POLLINICA.bat`.

### 🍎 macOS Users (Non-technical)

1. Download the project from GitHub (`Code → Download ZIP`) and extract it.
2. Open the `mac/` folder.
3. Read the `ISTRUZIONI_MAC.txt` file.
4. Launch the application via `AVVIA_CONTA_POLLINICA_GUI.sh`.

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

---

## ⚙️ Configuration and Conventions

**`.cfg` File** — The `codice/pollencounter.cfg` file allows you to modify work folders and calculation parameters without touching the source code.

**File Naming Convention** — It is essential to maintain the original names and folder structure. Files in the results folder must follow the `DD-MM-YYYY` date format.

**Excel Templates** — Do not modify the structure of the `riepilogo_settimana` and `dati_grezzi` sheets. To update the visual formatting, use the dedicated script:

```bash
python3 script_aiuto/applica_formattazione.py
```

---

## 🛠 Troubleshooting (FAQ)

**The executable won't start?** Ensure you have correctly extracted the ZIP archive and that your antivirus is not blocking the executable file.

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
| `pyinstaller` | Development only | Windows/macOS executable build |
