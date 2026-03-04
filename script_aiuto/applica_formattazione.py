"""
applica_formattazione.py — script one-shot.

Applica la formattazione visiva di "Tabelle monitoraggi ultima settimana gennaio.xlsx"
al foglio `riepilogo_settimana` di Polline_Template_Settimanale.xlsx
(file principale + copie in linux/ e windows/).

NON tocca: foglio 00_CODICI_SPECIE, foglio dati_grezzi, struttura dati.
"""

import shutil
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment

# ---------------------------------------------------------------------------
# Configurazione
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent.parent / "codice"
TEMPLATE_NAME = "Polline_Template_Settimanale.xlsx"

TEMPLATE_PRINCIPALE = BASE_DIR / TEMPLATE_NAME
COPIE = [
    BASE_DIR.parent / "windows" / TEMPLATE_NAME,
]

# ---------------------------------------------------------------------------
# Stili
# ---------------------------------------------------------------------------

VERDE = PatternFill(fill_type="solid", fgColor="C5E0B4")
NESSUN_FILL = PatternFill(fill_type=None)

thin = Side(border_style="thin", color="000000")
BORDO_THIN = Border(left=thin, right=thin, top=thin, bottom=thin)

FONT_BOLD = Font(bold=True)
FONT_NORMALE = Font(bold=False)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)

# Altezza riga in pt (unità openpyxl = pt)
ALTEZZA_RIGHE = 15


# ---------------------------------------------------------------------------
# Funzione principale
# ---------------------------------------------------------------------------

def applica_formattazione(path: Path) -> None:
    wb = openpyxl.load_workbook(path)
    ws = wb["riepilogo_settimana"]

    # --- Altezza righe ---
    for r in list(range(5, 54)) + list(range(57, 71)):
        ws.row_dimensions[r].height = ALTEZZA_RIGHE

    # --- Sfondo celle di conteggio: verde solo sulle righe rilevanti, bianco sulle altre ---
    # Sinistra: colonne G–M (7–13) | Destra: colonne S–X (19–24)
    # Righe con sfondo verde (pollini di interesse + spore rilevanti):
    VERDE_RIGHE = {10, 17, 18, 19, 20, 22, 23, 24, 27, 28, 29, 38, 40, 41, 52, 58}
    BIANCO = PatternFill(fill_type="solid", fgColor="FFFFFF")

    conteggio_cols_sx = range(7, 14)   # G=7 … M=13
    conteggio_cols_dx = range(19, 25)  # S=19 … X=24
    tutte_righe_dati = list(range(6, 53)) + list(range(58, 70))

    for row in tutte_righe_dati:
        fill = VERDE if row in VERDE_RIGHE else BIANCO
        for col in conteggio_cols_sx:
            ws.cell(row=row, column=col).fill = fill
        for col in conteggio_cols_dx:
            ws.cell(row=row, column=col).fill = fill

    # --- Bordi thin sul perimetro e sulle celle interne ---
    # Sezione sinistra:  D–M  (col 4–13), righe 5–53 e 57–70
    # Sezione destra:    P–X  (col 16–24), righe 5–53 e 57–70
    bordi_sezioni = [
        (range(4, 14), list(range(5, 54)) + list(range(57, 71))),   # D–M
        (range(16, 25), list(range(5, 54)) + list(range(57, 71))),  # P–X
    ]
    for cols, righe in bordi_sezioni:
        for row in righe:
            for col in cols:
                ws.cell(row=row, column=col).border = BORDO_THIN

    # --- Grassetto colonne specie (D e P): solo righe specificate ---
    SPECIE_BOLD_RIGHE = {
        7, 8, 9, 10,
        12, 13, 14, 15, 16, 17, 18,
        21, 22,
        25, 26, 27, 28, 29,
        36, 37, 38, 39, 40,
        44,
        46, 47, 48,
        52, 58,
    }
    specie_cols = [4, 16]  # D e P
    tutte_specie_righe = list(range(6, 53)) + list(range(58, 70))
    for col in specie_cols:
        for row in tutte_specie_righe:
            c = ws.cell(row=row, column=col)
            bold = row in SPECIE_BOLD_RIGHE
            c.font = Font(bold=bold, name=c.font.name, size=c.font.size)

    # --- Grassetto su intestazioni e righe totali (tutte le colonne del range) ---
    # Righe 5 e 57: intestazioni giornate | Righe 53 e 70: totali
    bold_rows_config = {
        5:  list(range(4, 14)) + list(range(16, 25)),  # intestazione pollini
        53: list(range(4, 14)) + list(range(16, 25)),  # totale pollini
        57: list(range(4, 14)) + list(range(16, 25)),  # intestazione spore
        70: list(range(4, 14)) + list(range(16, 25)),  # totale spore
    }
    for row, cols in bold_rows_config.items():
        for col in cols:
            c = ws.cell(row=row, column=col)
            c.font = Font(bold=True, name=c.font.name, size=c.font.size)

    # --- Allineamento centrato su tutto il range dati ---
    # D–M (col 4–13) e P–X (col 16–24), righe 5–53 e 57–70
    allinea_cols = list(range(4, 14)) + list(range(16, 25))
    allinea_righe = list(range(5, 54)) + list(range(57, 71))
    for row in allinea_righe:
        for col in allinea_cols:
            ws.cell(row=row, column=col).alignment = ALIGN_CENTER

    wb.save(path)
    print(f"  Salvato: {path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    targets = [TEMPLATE_PRINCIPALE] + COPIE

    # Backup del template principale
    backup_path = TEMPLATE_PRINCIPALE.with_name(
        TEMPLATE_PRINCIPALE.stem + "_BACKUP.xlsx"
    )
    if not backup_path.exists():
        shutil.copy2(TEMPLATE_PRINCIPALE, backup_path)
        print(f"Backup creato: {backup_path}")
    else:
        print(f"Backup già esistente (non sovrascritto): {backup_path}")

    print("\nApplicazione formattazione...")
    for target in targets:
        if target.exists():
            applica_formattazione(target)
        else:
            print(f"  [SALTATO — non trovato]: {target}")

    print("\nFatto. Aprire il template con LibreOffice/Excel per la verifica visiva.")


if __name__ == "__main__":
    main()
