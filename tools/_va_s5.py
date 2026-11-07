"""Foglio «5 - Riepilogo e prove»: riepilogo delle sei strade, scarti dichiarati dal motore mensile, cali storici, controlli."""
from __future__ import annotations

from _va_base import *  # noqa: F401,F403
import _va_base as B
import _va_s2, _va_s3, _va_s4

HDR = 5
R0 = 6
NOMI = ["Fondo mondiale, a distribuzione", "Fondo mondiale, ad accumulo", "Azioni italiane", "Un titolo di Stato decennale tenuto",
        "Buoni del Tesoro a 12 mesi", "Casa da affittare"]
# valori del motore mensile del video (partenza 2006, 100.000 euro) e scarti attesi dal foglio annuale (analista/out/07_scarto_excel_vs_motore.csv)
MOTORE = [395015.76594719174, 412725.3305485811, 214262.40757338592, 144168.9874260997, 129737.54444878639, 141410.1765373252]
SCARTO_ATTESO = [0.0196603, 0.0074653, 0.0071145, -0.0018610, -0.0032133, -0.0052441]
SOGLIA = 0.025
INFL_ROW = R0 + 7
ANNI_ROW = INFL_ROW + 1
CTRL_SEC = ANNI_ROW + 12


def build(ws):
    testata(ws, 10)
    s = inp("start")
    section_header(ws, 3, "Riepilogo: le sei strade a ottobre 2026 (netto di tasse e costi)", span=10)
    note_cell(ws, 4, 1, "Stima annuale: il foglio non replica il motore mensile del video. Gli scarti dal motore sono qui sotto, dichiarati. Valori reali in euro di partenza.", span=10)
    intest(ws, HDR, 1, ["Strada", "Netto finale (euro nominali)", "In euro di partenza", "Rendimento annuo nominale", "Rendimento annuo reale",
                        "Motore mensile del video (dato storico)", "Scarto atteso (prova)", "Scarto del foglio", "Dentro la soglia?"], altezza=60)
    srcs = [f"={q(S2)}G{_va_s2.RES_ROWS['dist']}", f"={q(S2)}G{_va_s2.RES_ROWS['acc']}", f"={q(S2)}G{_va_s2.RES_ROWS['it']}",
            f"={q(S3)}B{_va_s3.LIQ['netto']}", f"={q(S2)}G{_va_s2.RES_ROWS['bot']}", f"={q(S4)}B{_va_s4.RES['netto']}"]
    for k, nome in enumerate(NOMI):
        r = R0 + k
        label_cell(ws, r, 1, nome)
        output_cell(ws, r, 2, srcs[k], EUR)
        output_cell(ws, r, 3, f"=B{r}/(1+$B${INFL_ROW})", EUR)
        output_cell(ws, r, 4, f"=(B{r}/{inp('cap')})^(1/$B${ANNI_ROW})-1", PCT2)
        output_cell(ws, r, 5, f"=(C{r}/{inp('cap')})^(1/$B${ANNI_ROW})-1", PCT2)
        input_cell(ws, r, 6, MOTORE[k], EUR)
        input_cell(ws, r, 7, SCARTO_ATTESO[k], PCT2)
        output_cell(ws, r, 8, f"=IF(AND({s}=2006,{inp('cap')}=100000),B{r}/F{r}-1,\"non confrontabile\")", PCT2)
        output_cell(ws, r, 9, f"=IF(ISNUMBER(H{r}),IF(ABS(H{r})<={SOGLIA},\"sì\",\"no\"),\"-\")", "@")
    label_cell(ws, INFL_ROW, 1, "Inflazione cumulata fino a ottobre 2026 (stima)")
    foi = f"{q(S1)}$K${B_2005}:$K${B_Y1}"
    output_cell(ws, INFL_ROW, 2, f"={q(S1)}$K${B_Y1}/INDEX({foi},{s}-2005)-1", PCT2)
    label_cell(ws, ANNI_ROW, 1, "Anni di investimento (fino a ottobre 2026)")
    output_cell(ws, ANNI_ROW, 2, f"=SUMIF({q(S1)}$A${B_Y0}:$A${B_Y1},\">=\"&{s},{q(S1)}$B${B_Y0}:$B${B_Y1})", DEC4)
    note_cell(ws, ANNI_ROW + 1, 1, "Soglia di scarto dal motore mensile: 2,5% (fondo a distribuzione: +1,97%, perché il passo annuale reinveste i dividendi a metà anno). "
              "La soglia vale solo con partenza 2006 e 100.000 euro.", span=10)

    # frasi di sintesi (FIXED: numeri con separatore delle migliaia, senza TEXT)
    sec = ANNI_ROW + 3
    section_header(ws, sec, "In una riga per strada", span=10)
    for k, nome in enumerate(NOMI):
        r = R0 + k
        c = ws.cell(row=sec + 1 + k, column=1,
                    value=f"=\"{nome}: \"&FIXED(B{r},0)&\" € nominali, \"&FIXED(C{r},0)&\" € in euro di partenza\"")
        c.fill = FILL_OUTPUT
        c.font = FONT_VALUE
    ctrl = CTRL_SEC

    # cali storici
    st = ctrl + 12
    section_header(ws, st - 1, "Cali storici, nel momento peggiore (dati storici, non calcolati dal foglio)", span=9)
    intest(ws, st, 1, ["Strada", "Valuta", "Evento", "Picco", "Minimo", "Calo massimo (%)", "Recupero", "Mesi dal minimo al recupero", "Mesi dal picco al recupero"], altezza=45)
    nomi_chiari = {"ETF distribuzione": "Fondo mondiale, a distribuzione", "ETF accumulo": "Fondo mondiale, ad accumulo", "BTP": "Titolo di Stato decennale",
                   "BOT": "Buoni del Tesoro a 12 mesi"}
    for k, d in enumerate(B.leggi_csv("04_drawdown_eventi.csv")):
        r = st + 1 + k
        input_cell(ws, r, 1, nomi_chiari.get(d["strada"], d["strada"]), "@")
        input_cell(ws, r, 2, d["valuta"], "@")
        input_cell(ws, r, 3, d["evento"], "@")
        input_cell(ws, r, 4, d["picco"], "@")
        input_cell(ws, r, 5, d["minimo"], "@")
        input_cell(ws, r, 6, float(d["calo_max_pct"]) / 100, "0.0%")
        input_cell(ws, r, 7, d["recupero"] or "da recuperare", "@")
        for col, k_ in ((8, "mesi_minimo_recupero"), (9, "mesi_picco_recupero")):
            input_cell(ws, r, col, int(float(d[k_])) if d[k_] else "-", INT)
    st_end = st + 48

    # controlli
    section_header(ws, ctrl - 1, "Controlli", span=9)
    s3r0, s3r1 = _va_s3.R0, _va_s3.R0 + NY - 1
    checks = [
        ("Anno di partenza tra 2006 e 2016", f"=IF(AND({s}>=2006,{s}<=2016),\"OK\",\"controlla le celle gialle\")"),
        ("Capitale iniziale positivo", f"=IF({inp('cap')}>0,\"OK\",\"controlla le celle gialle\")"),
        ("Scarto massimo dal motore mensile entro 2,5% (solo partenza 2006, 100.000 euro)",
         f"=IF(AND({s}=2006,{inp('cap')}=100000),IF(MAX(MAX(H{R0}:H{R0+5}),-MIN(H{R0}:H{R0+5}))<={SOGLIA},\"OK\",\"scarto oltre soglia\"),\"non confrontabile\")"),
        ("Scarto massimo del valore del titolo di Stato dal motore entro 1%",
         f"=IF(MAX(MAX({q(S3)}AE{s3r0}:AE{s3r1}),-MIN({q(S3)}AE{s3r0}:AE{s3r1}))<=0.01,\"OK\",\"scarto oltre soglia\")"),
        ("Nessun errore di formula nel foglio Borsa", f"=IF(ISERROR(SUM({q(S2)}A1:N125)),\"errore\",\"OK\")"),
        ("Nessun errore di formula nei fogli titolo di Stato e casa", f"=IF(ISERROR(SUM({q(S3)}A1:AE80)+SUM({q(S4)}A1:R60)),\"errore\",\"OK\")"),
        ("Il netto del fondo a distribuzione è sotto quello ad accumulo (imposta sui dividendi anno per anno)",
         f"=IF(B{R0}<B{R0+1},\"OK\",\"verifica\")"),
        ("Plusvalenza della casa esente (più di cinque anni)", f"=IF({q(S4)}B{_va_s4.RES['plus']}=0,\"OK\",\"imposta dovuta\")"),
    ]
    for k, (lab, fx) in enumerate(checks):
        r = ctrl + k
        label_cell(ws, r, 1, lab)
        output_cell(ws, r, 2, fx, "@")
    note_cell(ws, ctrl + len(checks), 1, "Il motore mensile del video resta il riferimento. Questo foglio è una stima annuale per capire la meccanica: valori diversi di poco, mai di più della soglia dichiarata.", span=9)
    B.autofit(ws, max_w=34)
    ws.column_dimensions["A"].width = 60
    ws.freeze_panes = "B3"
