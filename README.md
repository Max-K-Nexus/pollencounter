# 🌿 Pollencounter

**Pollencounter** is a tool for pollen counting, designed to support aerobiological monitoring and the production of official bulletins.

The main goal is to reduce manual work on Excel files, standardize calculations, and make the management of weekly readings and reports easier. You do not need to be a programmer to use the Windows version.

---

## ✨ Main Features
* ✅ **Automation**: Drastically reduces manual data entry and repetitive calculations.
* ✅ **Standardization**: Uniform calculations to ensure the quality of aerobiological data.
* ✅ **Versatility**: Can be used via a Graphical User Interface (GUI) or Python scripts.
* ✅ **Ready to Use**: Includes a standalone Windows executable, requiring no Python installation.

---

## 📂 Repository Structure

```text
pollencounter/
├── codice/                 # Core software and configuration
│   ├── pollencounter.cfg   # Path settings and parameters
│   ├── polline_counter.py  # Processing logic (command line)
│   ├── polline_counter_gui.py # Version with Graphical User Interface
│   ├── Polline_Template_Settimanale.xlsx # Base template for calculations
│   └── concentrazioni_polliniche.xlsx     # Template for concentrations
├── letture_settimanali/    # INPUT folder (weekly count files)
│   ├── Conta_Pollinica_16-02-2026.xlsx
│   └── Conta_Pollinica_23-02-2026.xlsx
├── riferimenti/            # Historical tables and reference data
├── script_aiuto/           # Utilities for formatting and quick start (Bash/Python)
├── windows/                # Specific resources for Windows users
│   ├── Conta_Pollinica.exe # Ready-to-use executable
│   ├── AVVIA_CONTA_POLLINICA.bat # Quick start script
│   ├── ISTRUZIONI_WINDOWS.txt     # Specific guide for Windows
│   └── build_exe.bat       # Script to compile the executable (dev)
├── esempio di bollettino.pdf # Example of the final output
├── CHANGELOG.md            # Modification history
└── ISTRUZIONI.txt          # General documentation
```

## 🔄 Typical Workflow

* Data Collection: The operator fills in the weekly count Excel files in the letture_settimanali/ folder, following the naming convention Conta_Pollinica_DD-MM-YYYY.xlsx.

* Processing: The user launches the application via the Windows executable (Conta_Pollinica.exe) or via Python scripts. The program reads the input files, templates, and configuration in the codice/ folder.

* Output: The tool generates updated Excel files and reports ready for drafting the bulletin (as shown in esempio di bollettino.pdf).

## 🚀 Usage Guide
🪟 Windows Users (Non-technical)

This mode does not require Python installation.

 * 1 Download the project from GitHub (Code -> Download ZIP) and extract it.

 * 2 Open the windows/ folder.

 * 3 Read the ISTRUZIONI_WINDOWS.txt file.

    Launch the application by double-clicking on Conta_Pollinica.exe or AVVIA_CONTA_POLLINICA.bat.

## 🐍 Python Users (Developers)

For those who want to modify the code or integrate the script into other workflows.

Clone the repository:
    

    git clone [https://github.com/Max-K-Nexus/pollencounter.git](https://github.com/Max-K-Nexus/pollencounter.git)
    cd pollencounter

Install dependencies (make sure you have pandas and openpyxl installed):
    

    pip install pandas openpyxl

Run the script:


    python codice/polline_counter_gui.py

## ⚙️ Configuration and Conventions

.cfg File: The codice/pollencounter.cfg file allows you to modify folder paths and calculation parameters without editing the source code.

Naming Convention: It is crucial to keep the original file names and folder structure. Files in the letture_settimanali/ folder must strictly follow the indicated date format (DD-MM-YYYY).

## 🛠️ Troubleshooting (FAQ)

The executable won't start? Make sure you have extracted the ZIP archive correctly and that your antivirus is not blocking the executable file.

Errors with Excel files? Ensure you haven't modified or moved the column structure in the templates located in the codice/ folder.

Using Mac/Linux? The .exe file is for Windows only. On other operating systems, you must use the Python scripts in the codice/ folder.

## 👥 Authors and Contributions

The Pollencounter project was developed by:

* Simone Bettella — Concept, original development, and calculation logic.

* Massimiliano Iotti — Maintenance, automation, documentation, and Windows support.

## 📜 License

The project is distributed under the GNU General Public License v3.0 (GPL‑3.0).
If you use Pollencounter in your project or report, please kindly cite the authors:

    Pollencounter – developed by Simone Bettella and Massimiliano Iotti

    https://github.com/Max-K-Nexus/pollencounter
