"""Base del builder venti_anni_etf_btp_casa_2026: workbook, stili, nomi dei fogli, posizioni delle celle, dati storici.
Specifica: _lab_adversarial_2026_08/venti_anni_etf_btp_casa/11_ANALISI_v2.md sez. 5 (sei fogli). Dati: analista/out/07_dati_excel.csv, 07b_dati_btp_per_partenza.csv.
Moduli: _va_base (questo), _va_s1 (dati e fisco), _va_s2 (borsa), _va_s3 (titolo di Stato), _va_s4 (casa), _va_s5 (riepilogo e prove)."""
from __future__ import annotations

import csv
import os
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _excel_helpers as H  # noqa: E402
from _excel_helpers import (  # noqa: E402,F401
    FONT_VALUE, FONT_LABEL, FONT_NOTE, BORDER, FILL_OUTPUT, FILL_INPUT, FILL_TABLE_HEADER, FILL_HEADER, FONT_HEADER,
    ALIGN_LEFT, ALIGN_CENTER, title_row, disclaimer_row, build_cover, section_header, label_cell, input_cell, output_cell,
    table_header, note_cell,
)

OUT = Path(__file__).resolve().parent.parent / "simulatori" / "venti_anni_etf_btp_casa_2026.xlsx"
DOSSIER = Path(os.environ.get(
    "VA_DOSSIER", "E:/sviluppo/confyui/virtual_influencer/_lab_adversarial_2026_08/venti_anni_etf_btp_casa"))
DATI = DOSSIER / "analista" / "out"

EUR = '#,##0" €"'
EUR2 = '#,##0.00" €"'
PCT2 = "0.00%"
PCT3 = "0.000%"
INT = "0"
DEC2 = "0.00"
DEC4 = "0.0000"
DATAF = "dd/mm/yyyy"
S0, S1, S2, S3, S4, S5 = ("0 - Indice", "1 - Dati e fisco", "2 - Borsa", "3 - BTP", "4 - Casa", "5 - Riepilogo e prove")
TITOLO = "20 anni: fondo mondiale, titolo di Stato, casa (caso inventato, numeri storici reali)"
DISC = ("Caso inventato e illustrativo: i tre fratelli non esistono, i numeri storici sì. Non è consulenza e non è una previsione. "
        "Stima annuale (il video usa un motore mensile): scarti dichiarati nel foglio 5. Ottobre 2026 è una stima.")
wb = Workbook()

ANNI = list(range(2006, 2027))
NY = len(ANNI)            # 21 anni
PARTENZE = list(range(2006, 2017))

# ---- foglio 1: posizioni fisse (input utente in colonna C) ----
INP = dict(start=4, cap=5, ter=6, fee_pct=7, fee_min=8, costi_acq=9, costi_vend=10, sfitto=11, manut=12, catasto=13,
           rend0=14, rend1=15, modo=16, cedolare=17, irpef06=18, irpef0710=19, addiz=20, extra=21,
           interp=22, aliq_ts=23, aliq_cg=24, aliq_cg_old=25, ded=26, reg=27)
# dati storici annuali: riga 2005 (solo indice prezzi) e righe 2006..2026
B_HDR = 31
B_2005 = 32
B_Y0 = 33                  # riga 2006
B_Y1 = B_Y0 + NY - 1       # riga 2026
# matrice rendimento fine anno (partenze in colonna B..L)
M_SEC = B_Y1 + 3
M_HDR = M_SEC + 1
M_Y0 = M_HDR + 1
M_Y1 = M_Y0 + NY - 1
# titolo comprato per partenza
T_SEC = M_Y1 + 3
T_HDR = T_SEC + 1
T_P0 = T_HDR + 1
T_P1 = T_P0 + len(PARTENZE) - 1
# tabella fiscale per anno
F_SEC = T_P1 + 3
F_HDR = F_SEC + 1
F_Y0 = F_HDR + 1
F_Y1 = F_Y0 + NY - 1
# scaglioni del bollo 2011
G_SEC = F_Y1 + 3
G_HDR = G_SEC + 1
G_R0 = G_HDR + 1
G_R1 = G_R0 + 3

# colonne dei dati storici (foglio 1)
DC = dict(anno="A", f="B", w_net="C", w_prc="D", i_gr="E", i_prc="F", casa="G", hicp="H", foi="I", bot="J", foi_idx="K")
# colonne tabella fiscale
FC = dict(anno="A", aliq="B", ts="C", b_fix="D", b_peso="E", b_pct="F", b_min="G", b_max="H", tobin="I", ici="J", molt="K", norma="L")


def q(sheet):
    return "'" + sheet + "'!"


def inp(key):
    """Riferimento assoluto a un input del foglio 1."""
    return f"{q(S1)}$C${INP[key]}"


def drow(i):
    """Riga (foglio 1) dei dati storici dell'anno i-esimo (0 = 2006)."""
    return B_Y0 + i


def frow(i):
    return F_Y0 + i


def dref(key, i):
    return f"{q(S1)}${DC[key]}${drow(i)}"


def fref(key, i):
    return f"{q(S1)}${FC[key]}${frow(i)}"


def fee_expr(x):
    """Commissione di negoziazione: % con minimo, mai oltre l'importo."""
    return f"MIN({x},MAX({inp('fee_min')},{inp('fee_pct')}*{x}))"


def bollo_expr(i, base, f):
    """Imposta di bollo dell'anno i-esimo sul valore base, per la frazione d'anno f (tabella fiscale del foglio 1)."""
    tab = f"{q(S1)}$A${G_R0}:$B${G_R1}"
    return (f"{f}*({fref('b_fix', i)}+{fref('b_peso', i)}*VLOOKUP({base},{tab},2,TRUE)"
            f"+IF({fref('b_pct', i)}>0,MIN({fref('b_max', i)},MAX({fref('b_min', i)},{fref('b_pct', i)}*({base}))),0))")


# ---- dati ----
def leggi_csv(nome):
    with open(DATI / nome, encoding="utf-8", newline="") as fh:
        return list(csv.DictReader(fh))


def dati_annuali():
    return {int(r["anno"]): r for r in leggi_csv("07_dati_excel.csv")}


def dati_btp():
    out = {}
    for r in leggi_csv("07b_dati_btp_per_partenza.csv"):
        out[(int(r["partenza"]), int(r["anno"]))] = r
    return out


# titolo comprato a inizio partenza e tassi a 10 anni nel mese dei rinnovi (dal motore: IRS nel mese di scadenza). Fonte: 07b + motore.run_btp
TITOLI = {
    2006: ("BTP 3,75% 1.2.2005 - 1.8.2015", 0.0375, date(2015, 8, 1), 102.49, [0.01839, 0.03538]),
    2007: ("BTP 4,0% 1.8.2006 - 1.2.2017", 0.04, date(2017, 2, 1), 98.95, [0.02352]),
    2008: ("BTP 4,50% 1.8.2007 - 1.2.2018", 0.045, date(2018, 2, 1), 99.41, [0.02081]),
    2009: ("BTP 4,50% 1.9.2008 - 1.3.2019", 0.045, date(2019, 3, 1), 100.95, [0.0269]),
    2010: ("BTP 4,25% 1.9.2009 - 1.3.2020", 0.0425, date(2020, 3, 1), 101.82, [0.01551]),
    2011: ("BTP 3,75% 1.9.2010 - 1.3.2021", 0.0375, date(2021, 3, 1), 92.09, [0.00704]),
    2012: ("BTP 5,0% 1.9.2011 - 1.3.2022", 0.05, date(2022, 3, 1), 86.53, [0.01854]),
    2013: ("BTP 5,50% 1.5.2012 - 1.11.2022", 0.055, date(2022, 11, 1), 108.43, [0.04243]),
    2014: ("BTP 4,50% 1.8.2013 - 1.3.2024", 0.045, date(2024, 3, 1), 103.53, [0.03696]),
    2015: ("BTP 2,50% 1.9.2014 - 1.12.2024", 0.025, date(2024, 12, 1), 105.58, [0.03321]),
    2016: ("BTP 2% 1.9.2015 - 1.12.2025", 0.02, date(2025, 12, 1), 103.81, [0.03552]),
}


def autofit(ws, min_w=9, max_w=40, skip_rows=(1, 2)):
    """Larghezza colonne dal contenuto (celle unite e righe titolo escluse)."""
    merged = set()
    for rng in ws.merged_cells.ranges:
        for row in ws.iter_rows(min_row=rng.min_row, max_row=rng.max_row, min_col=rng.min_col, max_col=rng.max_col):
            for c in row:
                merged.add(c.coordinate)
    widths = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None or c.coordinate in merged or c.row in skip_rows:
                continue
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                n = 14
            elif isinstance(v, float):
                n = 13
            else:
                n = len(str(v)) + 2
            widths[c.column_letter] = max(widths.get(c.column_letter, 0), n)
    for col, w in widths.items():
        ws.column_dimensions[col].width = max(min_w, min(max_w, w))


def intest(ws, row, col, testi, altezza=45):
    """Intestazioni di tabella azzurre, a capo."""
    for k, t in enumerate(testi):
        table_header(ws, row, col + k, t)
    ws.row_dimensions[row].height = altezza


def testata(ws, span, note=None):
    title_row(ws, 1, TITOLO, span=span)
    ws.row_dimensions[1].height = 32
    disclaimer_row(ws, 2, span=span)
    ws["A2"] = DISC
    ws.row_dimensions[2].height = 48
