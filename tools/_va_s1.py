"""Foglio «1 - Dati e fisco»: input utente, dati storici annuali, rendimenti di fine anno del titolo, titolo comprato, tabella fiscale."""
from __future__ import annotations

from _va_base import *  # noqa: F401,F403
import _va_base as B


def _inp(ws, key, label, val, fmt, nota):
    r = INP[key]
    label_cell(ws, r, 1, label)
    ws.cell(row=r, column=2, value="").border = BORDER
    input_cell(ws, r, 3, val, fmt)
    note_cell(ws, r, 4, nota, span=6)


def build(ws):
    testata(ws, 12)
    section_header(ws, 3, "A. Input: celle gialle, cambia solo queste", span=9)
    _inp(ws, "start", "Anno di partenza (2006-2016)", 2006, INT, "Capitale versato a fine dicembre dell'anno precedente. Il video usa il 2006.")
    _inp(ws, "cap", "Capitale iniziale", 100000, EUR, "Eredità del caso inventato.")
    _inp(ws, "ter", "Costo annuo del fondo (TER)", 0.005, PCT2, "Ipotesi dichiarata: 0,50% l'anno, come il fondo mondiale a distribuzione del 2006.")
    _inp(ws, "fee_pct", "Commissione di negoziazione", 0.001, PCT2, "Ipotesi: percentuale sul controvalore di ogni acquisto o vendita.")
    _inp(ws, "fee_min", "Commissione minima", 5, EUR2, "Ipotesi: minimo per ordine.")
    _inp(ws, "costi_acq", "Costi di acquisto della casa", 0.08, PCT2, "Ipotesi: imposte, notaio, agenzia. Prezzo = capitale / (1 + costi).")
    _inp(ws, "costi_vend", "Costi di vendita della casa", 0.03, PCT2, "Ipotesi: agenzia. Plusvalenza esente oltre cinque anni (L. 266/2005 c. 496).")
    _inp(ws, "sfitto", "Mesi senza inquilino (frazione d'anno)", 1 / 12, PCT2, "Ipotesi: un mese all'anno vuoto = 8,33%.")
    _inp(ws, "manut", "Manutenzione e spese non ripetibili", 0.01, PCT2, "Ipotesi: percentuale del valore della casa, ogni anno.")
    _inp(ws, "catasto", "Valore catastale / prezzo di acquisto", 0.35, PCT2, "Ipotesi: base fissa in euro per la tassa sulla casa (ICI fino al 2011, IMU dopo).")
    _inp(ws, "rend0", "Rendimento lordo dell'affitto all'avvio (2006)", 0.051, PCT2, "Nomisma, 13 grandi aree urbane, 2006 (fonte secondaria).")
    _inp(ws, "rend1", "Rendimento lordo a ottobre 2026", 0.0438, PCT2, "Media Italia abitazioni civili, primo semestre 2026 (fonte secondaria).")
    ws.cell(row=INP["modo"], column=1)
    label_cell(ws, INP["modo"], 1, "Come si calcola il canone")
    ws.cell(row=INP["modo"], column=2, value="").border = BORDER
    c = input_cell(ws, INP["modo"], 3, "rendimento", "@")
    note_cell(ws, INP["modo"], 4, "Scrivi rendimento (rendimento lordo sul valore) oppure HICP (canone iniziale indicizzato all'indice europeo degli affitti, stima annuale).", span=6)
    _inp(ws, "cedolare", "Tassa fissa sugli affitti (cedolare secca)", 0.21, PCT2, "Dal 2011 (D.Lgs. 23/2011 art. 3), aliquota ordinaria 21%.")
    _inp(ws, "irpef06", "Aliquota IRPEF di scaglione, 2006", 0.39, PCT2, "Ipotesi: reddito 35.000 euro più canone. L. 311/2004.")
    _inp(ws, "irpef0710", "Aliquota IRPEF di scaglione, 2007-2010", 0.38, PCT2, "Ipotesi: stesso reddito. L. 296/2006.")
    _inp(ws, "addiz", "Addizionale regionale", 0.009, PCT2, "Ipotesi: aliquota base 0,9%.")
    _inp(ws, "extra", "Punti di IRPEF in più (prova)", 0.0, PCT2, "0 = caso base. Sale la tassa sugli affitti fino al 2010.")
    _inp(ws, "interp", "Anni su cui si interpola il rendimento", 20.33, DEC2, "Dal 2006 a ottobre 2026: interpolazione lineare.")
    _inp(ws, "aliq_ts", "Imposta sui titoli di Stato", 0.125, PCT2, "Aliquota sostitutiva su cedole e guadagni dei titoli di Stato (D.Lgs. 239/1996).")
    _inp(ws, "aliq_cg", "Imposta sui guadagni dei fondi dal 2015", 0.26, PCT2, "Aliquota sostitutiva 26% (DL 66/2014).")
    _inp(ws, "aliq_cg_old", "Imposta sui guadagni fino al 30/06/2014", 0.20, PCT2, "Aliquota 20% (DL 138/2011): usata per la correzione sul guadagno maturato al 30/06/2014.")
    _inp(ws, "ded", "Quota del canone che entra nel reddito (fino al 2010)", 0.85, PCT2, "Deduzione forfettaria del 15% sul canone (TUIR art. 37).")
    _inp(ws, "reg", "Imposta di registro a carico del proprietario (fino al 2010)", 0.01, PCT2, "1% del canone annuo, quota del locatore (ipotesi).")
    note_cell(ws, 28, 1, "La lettura (b), redditi incassati e spesi, non è nel foglio: è solo nel video. Nel foglio il capitale resta investito e si liquida in fondo.", span=9)

    # ---- B. dati storici annuali ----
    section_header(ws, B_HDR - 1, "B. Dati storici annuali: celle gialle, la fonte è nelle note in fondo", span=11)
    intest(ws, B_HDR, 1, ["Anno", "Frazione d'anno", "Fondo mondiale, con dividendi", "Fondo mondiale, solo prezzo",
                          "Azioni italiane, con dividendi", "Azioni italiane, solo prezzo", "Prezzi delle case",
                          "Affitti (indice europeo)", "Inflazione (indice FOI)", "Buoni del Tesoro a 12 mesi (media)",
                          "Indice FOI a fine anno"])
    D = B.dati_annuali()
    foi05 = 70.80789
    ws.cell(row=B_2005, column=1, value=2005).font = FONT_LABEL
    input_cell(ws, B_2005, 11, foi05, "0.00000")
    for i, y in enumerate(ANNI):
        r = drow(i)
        d = D[y]
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        input_cell(ws, r, 2, float(d["frazione_anno"]) if y < 2026 else 10 / 12, "0.0000")
        for col, k in ((3, "World_NETR_pct"), (4, "World_prezzo_pct"), (5, "Italia_GRTR_pct"), (6, "Italia_prezzo_pct"),
                       (7, "Casa_prezzo_pct"), (8, "Affitti_HICP_pct"), (9, "FOI_pct"), (10, "BOT12m_medio_pct")):
            input_cell(ws, r, col, float(d[k]) / 100, PCT2)
        input_cell(ws, r, 11, float(d["FOI_indice_fine"]), "0.00000")
    note_cell(ws, B_Y1 + 1, 1, "Fonti: indice azionario mondiale e indice azionario italiano (con dividendi, in euro), indice nazionale dei prezzi delle case, indice europeo degli affitti, FOI Istat, "
              "rendimento dei Buoni del Tesoro a 12 mesi (Banca d'Italia). Ultimo anno: gennaio-ottobre 2026 (dieci mesi su dodici), stima fino al dato definitivo.", span=11)

    # ---- matrice rendimento di fine anno del titolo detenuto ----
    section_header(ws, M_SEC, "C. Rendimento stimato a fine anno del titolo detenuto, per anno di partenza (stima, non calcolata dal foglio)", span=12)
    intest(ws, M_HDR, 1, ["Anno \\ partenza"] + PARTENZE, altezza=30)
    BT = B.dati_btp()
    for i, y in enumerate(ANNI):
        r = M_Y0 + i
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        for k, p in enumerate(PARTENZE):
            row = BT.get((p, y))
            if row is not None:
                input_cell(ws, r, 2 + k, float(row["rendimento_fine"]), PCT3)
            else:
                cc = ws.cell(row=r, column=2 + k, value=None)
                cc.border = BORDER
    note_cell(ws, M_Y1 + 1, 1, "Stima storica (curva dei Buoni del Tesoro e tassi a 10 anni, aste Banca d'Italia): serve alla funzione PREZZO() del foglio 3. Celle vuote = anni prima della partenza.", span=12)

    # ---- titolo comprato per partenza ----
    section_header(ws, T_SEC, "D. Titolo di Stato decennale comprato a inizio anno e tassi dei rinnovi (stima storica)", span=9)
    intest(ws, T_HDR, 1, ["Partenza", "Titolo (descrizione)", "Cedola annua", "Scadenza", "Prezzo di acquisto", "Data di regolamento",
                          "Tasso a 10 anni al 1° rinnovo", "Tasso a 10 anni al 2° rinnovo"])
    for k, p in enumerate(PARTENZE):
        r = T_P0 + k
        desc, ced, scad, pr, rin = B.TITOLI[p]
        ws.cell(row=r, column=1, value=p).font = FONT_LABEL
        input_cell(ws, r, 2, desc, "@")
        input_cell(ws, r, 3, ced, PCT2)
        input_cell(ws, r, 4, scad, DATAF)
        input_cell(ws, r, 5, pr, DEC2)
        input_cell(ws, r, 6, date(p, 1, 3), DATAF)
        input_cell(ws, r, 7, rin[0], PCT3)
        if len(rin) > 1:
            input_cell(ws, r, 8, rin[1], PCT3)
        else:
            input_cell(ws, r, 8, 0, PCT3)
    note_cell(ws, T_P1 + 1, 1, "Il primo titolo è reale (asta di Banca d'Italia di dicembre 2005). Alla scadenza si rinnova con un decennale tipo, non un titolo reale: cedola = tasso a 10 anni del mese, "
              "arrotondato al quarto di punto. Valori storici, non calcolati dal foglio.", span=9)

    # ---- tabella fiscale ----
    section_header(ws, F_SEC, "E. Tabella fiscale per anno: ogni riga ha la norma e il grado della fonte (P = primaria)", span=12)
    intest(ws, F_HDR, 1, ["Anno", "Imposta su dividendi e guadagni dei fondi", "Imposta sui titoli di Stato", "Bollo fisso (quota a 34,20 €)",
                          "Peso del bollo a scaglioni", "Bollo in % del valore", "Bollo minimo", "Bollo massimo", "Imposta sulle transazioni (azioni)",
                          "Aliquota tassa sulla casa", "Moltiplicatore catastale", "Norma e fonte (grado)"], altezza=60)
    for i, y in enumerate(ANNI):
        r = frow(i)
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        aliq = 0.125 if y <= 2011 else 0.20 if y <= 2013 else 0.23 if y == 2014 else 0.26
        fix = 34.20 if y <= 2010 else 197 / 365 * 34.20 if y == 2011 else 0.0
        peso = 168 / 365 if y == 2011 else 0.0
        pct = 0.001 if y == 2012 else 0.0015 if y == 2013 else 0.002 if y >= 2014 else 0.0
        bmin = 34.20 if y in (2012, 2013) else 0.0
        bmax = 1200.0 if y == 2012 else (1e9 if y >= 2013 else 0.0)
        tob = 0.0 if y <= 2012 else 0.0012 if y == 2013 else 0.0020 if y == 2026 else 0.0010
        ici = 0.0064 if y == 2006 else 0.00646 if y == 2007 else 0.00649 if y <= 2011 else 0.00933
        molt = 1.0 if y <= 2011 else 1.6
        input_cell(ws, r, 2, aliq, PCT2)
        input_cell(ws, r, 3, 0.125, PCT2)
        input_cell(ws, r, 4, fix, EUR2)
        input_cell(ws, r, 5, peso, "0.0000")
        input_cell(ws, r, 6, pct, PCT2)
        input_cell(ws, r, 7, bmin, EUR2)
        input_cell(ws, r, 8, bmax, EUR2)
        input_cell(ws, r, 9, tob, PCT2)
        input_cell(ws, r, 10, ici, "0.000%")
        input_cell(ws, r, 11, molt, "0.0")
        if y <= 2010:
            n = "Rendite 12,5% (D.Lgs. 461/1997); bollo 34,20 € fissi (DPR 642/1972, circ. AE 40/E/2011); ICI media nazionale IFEL (P)"
        elif y == 2011:
            n = "Bollo 2011: 34,20 € fino al 16/07, poi 70 € sullo scaglione 50-150 mila € (DL 98/2011); dal 2012 uno per mille (DL 201/2011); IMU dal 2012 (P)"
        elif y == 2012:
            n = "Rendite 20% (DL 138/2011); bollo 1 per mille, min 34,20 max 1.200 € (DL 201/2011 art. 19); IMU 9,33 per mille, moltiplicatore 1,6 (P)"
        elif y == 2013:
            n = "Bollo 1,5 per mille, min 34,20 € (DL 201/2011); imposta sulle transazioni 0,12% (L. 228/2012) (P)"
        elif y == 2014:
            n = "Rendite 20% fino al 30/06, 26% da luglio: media 23%; bollo 2 per mille (L. 147/2013 c. 581); transazioni 0,10% (P)"
        elif y == 2026:
            n = "Transazioni: aliquota 0,20% nel 2026 (da verificare nella scheda numeri); resto invariato (P/E)"
        else:
            n = "Rendite 26% (DL 66/2014); bollo 2 per mille; transazioni 0,10% (P)"
        c = ws.cell(row=r, column=12, value=n)
        c.font = FONT_NOTE
        c.alignment = ALIGN_LEFT

    # ---- scaglioni del bollo 2011 ----
    section_header(ws, G_SEC, "F. Scaglioni del bollo 2011 (da luglio): valore del conto titoli", span=3)
    intest(ws, G_HDR, 1, ["Da (euro)", "Bollo annuo (euro)"], altezza=30)
    for k, (s, b) in enumerate(((0, 34.20), (50000, 70.0), (150000, 240.0), (500000, 680.0))):
        input_cell(ws, G_R0 + k, 1, s, EUR)
        input_cell(ws, G_R0 + k, 2, b, EUR2)
    note_cell(ws, G_R1 + 1, 1, "Soglie del DL 98/2011: sotto 50.000 € il bollo resta 34,20 €; nello scaglione 50-150 mila € (il caso dei tre fratelli) è 70 € l'anno; dal 2012 diventa uno per mille (DL 201/2011). Il peso 197/365 e 168/365 divide il 2011 prima e dopo il 17/07.", span=8)
    B.autofit(ws, max_w=30)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["L"].width = 60
    ws.freeze_panes = "A3"
