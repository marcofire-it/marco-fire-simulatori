"""Foglio «3 - BTP»: un titolo di Stato decennale tenuto, rinnovato alla scadenza; valore con PREZZO(), cedole, conto dei Buoni del Tesoro, liquidazione."""
from __future__ import annotations

from _va_base import *  # noqa: F401,F403
import _va_base as B

HDR = 25
R0 = 26                       # riga del primo anno (2006)
VM_SEC = R0 + NY + 3
VM_HDR = VM_SEC + 1
VM_Y0 = VM_HDR + 1
CELL = dict(desc=5, ced=6, scad=7, prezzo=8, regol=9, rateo=10, nom=11, t1=12, t2=13)
LIQ = dict(valore=16, conto=17, comm=18, imposta=19, netto=20)


def row(i):
    return R0 + i


def build(ws):
    testata(ws, 14)
    s = inp("start")
    t1 = lambda col: (f"INDEX({q(S1)}${col}${T_P0}:${col}${T_P1},MATCH({s},{q(S1)}$A${T_P0}:$A${T_P1},0))")  # noqa: E731
    section_header(ws, 3, "Il titolo di Stato comprato all'inizio (preso dalla tabella D del foglio 1)", span=6)
    voci = [("desc", "Titolo", f"={t1('B')}", "@"), ("ced", "Cedola annua", f"={t1('C')}", PCT2), ("scad", "Scadenza", f"={t1('D')}", DATAF),
            ("prezzo", "Prezzo di acquisto (su 100)", f"={t1('E')}", DEC2), ("regol", "Data di regolamento", f"={t1('F')}", DATAF)]
    for k, lab, fx, fm in voci:
        label_cell(ws, CELL[k], 1, lab)
        output_cell(ws, CELL[k], 2, fx, fm)
    label_cell(ws, CELL["rateo"], 1, "Rateo di cedola maturato all'acquisto (su 100)")
    output_cell(ws, CELL["rateo"], 2, f"=B{CELL['ced']}/2*100*COUPDAYBS(B{CELL['regol']},B{CELL['scad']},2,1)/COUPDAYS(B{CELL['regol']},B{CELL['scad']},2,1)", DEC4)
    label_cell(ws, CELL["nom"], 1, "Valore nominale comprato")
    output_cell(ws, CELL["nom"], 2, f"=({inp('cap')}-{B.fee_expr(inp('cap'))})*100/(B{CELL['prezzo']}+B{CELL['rateo']})", EUR)
    label_cell(ws, CELL["t1"], 1, "Tasso a 10 anni al primo rinnovo")
    output_cell(ws, CELL["t1"], 2, f"={t1('G')}", PCT3)
    label_cell(ws, CELL["t2"], 1, "Tasso a 10 anni al secondo rinnovo")
    output_cell(ws, CELL["t2"], 2, f"={t1('H')}", PCT3)
    note_cell(ws, 14, 1, "Un solo titolo, non un paniere: alla scadenza si rinnova con un decennale tipo (non un titolo reale), con cedola uguale al rendimento di allora, arrotondato al quarto di punto.", span=6)

    section_header(ws, 15, "Se vendesse a ottobre 2026 (euro nominali, netto di tasse e costi)", span=6)
    last = row(NY - 1)
    lab = {"valore": "Valore del titolo (prezzo di mercato più rateo)", "conto": "Conto dei Buoni del Tesoro (cedole reinvestite)",
           "comm": "Commissione di vendita", "imposta": "Imposta sul guadagno di prezzo", "netto": "Netto finale"}
    vv, vc, vm, vi = LIQ["valore"], LIQ["conto"], LIQ["comm"], LIQ["imposta"]
    fx = {"valore": f"=Z{last}", "conto": f"=AC{last}", "comm": "=" + B.fee_expr(f"B{vv}"),
          "imposta": f"={inp('aliq_ts')}*MAX(0,(X{last}-R{last})/100*O{last})",
          "netto": f"=B{vv}+B{vc}-B{vm}-B{vi}"}
    for k in LIQ:
        label_cell(ws, LIQ[k], 1, lab[k])
        output_cell(ws, LIQ[k], 2, fx[k], EUR)
    note_cell(ws, 21, 1, "Il valore NON è incollato dal motore: viene da PREZZO() applicato al rendimento stimato di fine anno (foglio 1, tabella C). "
              "Scarto dal motore mensile (partenza 2006): -0,19% sul netto; massimo 0,54% sul valore a fine anno (colonna di prova). "
              "PREZZO() non accetta rendimenti negativi: in quegli anni il foglio usa la stessa formula scritta per esteso.", span=6)

    section_header(ws, HDR - 1, "Motore del calcolo, non modificare: anno per anno", span=31)
    intest(ws, HDR, 1, ["Anno", "In corso (1 = sì)", "Frazione d'anno", "Scadenza a inizio anno", "Cedola a inizio anno", "Nominale a inizio anno", "Prezzo di carico",
                        "Ultimo mese dell'anno", "Mese cedola 1", "Mese cedola 2", "Rinnovo nell'anno (1 = sì)", "Rinnovi già fatti", "Tasso a 10 anni al rinnovo",
                        "Rimborso netto di imposta", "Nominale a fine anno", "Cedola a fine anno", "Scadenza a fine anno", "Prezzo di carico a fine anno",
                        "Cedole pagate dal vecchio titolo", "Cedole pagate dal nuovo titolo", "Cedole lorde incassate", "Data di valutazione",
                        "Rendimento a fine anno (stima)", "Prezzo secco (PREZZO)", "Rateo", "Valore del titolo", "Conto prima del bollo", "Imposta di bollo",
                        "Conto a fine anno", "Valore del motore mensile (dato storico)", "Scarto del foglio"], altezza=75)
    mat = f"{q(S1)}$B${M_Y0}:$L${M_Y1}"
    mhd = f"{q(S1)}$B${M_HDR}:$L${M_HDR}"
    vmat = f"$B${VM_Y0}:$L${VM_Y0 + NY - 1}"
    vhd = f"$B${VM_HDR}:$L${VM_HDR}"
    ts = inp("aliq_ts")
    for i, y in enumerate(ANNI):
        r = row(i)
        p = r - 1
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        F = {
            2: (f"=IF(A{r}>={s},1,0)", INT),
            3: (f"={dref('f', i)}", "0.0000"),
            4: (f"=IF(B{r}=0,0,IF(A{r}={s},$B${CELL['scad']},Q{p}))", DATAF),
            5: (f"=IF(B{r}=0,0,IF(A{r}={s},$B${CELL['ced']},P{p}))", PCT2),
            6: (f"=IF(B{r}=0,0,IF(A{r}={s},$B${CELL['nom']},O{p}))", EUR),
            7: (f"=IF(B{r}=0,0,IF(A{r}={s},$B${CELL['prezzo']},R{p}))", DEC2),
            8: (f"=ROUND(C{r}*12,0)", INT),
            9: (f"=IF(B{r}=0,0,MONTH(D{r}))", INT),
            10: (f"=IF(B{r}=0,0,MOD(I{r}+5,12)+1)", INT),
            11: (f"=IF(B{r}=0,0,IF(AND(YEAR(D{r})=A{r},I{r}<=H{r}),1,0))", INT),
            12: (f"=IF(B{r}=0,0,IF(A{r}={s},0,L{p}+K{p}))", INT),
            13: (f"=IF(K{r}=1,INDEX($B${CELL['t1']}:$B${CELL['t2']},L{r}+1),0)", PCT3),
            14: (f"=IF(K{r}=1,F{r}-{ts}*MAX(0,(100-G{r})/100*F{r}),0)", EUR),
            15: (f"=IF(B{r}=0,0,IF(K{r}=1,N{r}-{B.fee_expr(f'N{r}')},F{r}))", EUR),
            16: (f"=IF(B{r}=0,0,IF(K{r}=1,MAX(ROUND(M{r}/0.0025,0)*0.0025,0.0025),E{r}))", PCT2),
            17: (f"=IF(B{r}=0,0,IF(K{r}=1,EDATE(D{r},120),D{r}))", DATAF),
            18: (f"=IF(B{r}=0,0,IF(K{r}=1,100,G{r}))", DEC2),
            19: (f"=IF(B{r}=0,0,IF(K{r}=1,1+IF(J{r}<I{r},1,0),IF(I{r}<=H{r},1,0)+IF(J{r}<=H{r},1,0)))", INT),
            20: (f"=IF(K{r}=1,IF(AND(J{r}>I{r},J{r}<=H{r}),1,0),0)", INT),
            21: (f"=IF(B{r}=0,0,F{r}*E{r}/2*S{r}+O{r}*P{r}/2*T{r})", EUR),
            22: (f"=IF(B{r}=0,0,MIN(DATE(A{r},12,31),DATE(2026,10,30)))", DATAF),
            23: (f"=IF(B{r}=0,0,INDEX({mat},A{r}-2005,MATCH({s},{mhd},0)))", PCT3),
            24: (f"=IF(B{r}=0,0,IF(W{r}>=0,PRICE(V{r},Q{r},P{r},W{r},100,2,1),"
                 f"(1+W{r}/2)^(1-COUPDAYSNC(V{r},Q{r},2,1)/COUPDAYS(V{r},Q{r},2,1))*(P{r}/2*100*(-PV(W{r}/2,COUPNUM(V{r},Q{r},2,1),1))"
                 f"+100*(1+W{r}/2)^(-COUPNUM(V{r},Q{r},2,1)))-Y{r}))", "0.0000"),
            25: (f"=IF(B{r}=0,0,P{r}/2*100*COUPDAYBS(V{r},Q{r},2,1)/COUPDAYS(V{r},Q{r},2,1))", "0.0000"),
            26: (f"=IF(B{r}=0,0,O{r}*(X{r}+Y{r})/100)", EUR),
            27: (f"=IF(B{r}=0,0,IF(A{r}={s},0,AC{p})*(1+{dref('bot', i)}*(1-{ts})*C{r})+U{r}*(1-{ts}))", EUR),
            28: (f"=IF(B{r}=0,0,{B.bollo_expr(i, f'(AA{r}+Z{r})', f'C{r}')})", EUR),
            29: (f"=IF(B{r}=0,0,AA{r}-AB{r})", EUR),
        }
        for col, (fx_, fm) in F.items():
            output_cell(ws, r, col, fx_, fm)
        # dato storico del motore (matrice sotto) e scarto
        output_cell(ws, r, 30, f"=IF(B{r}=0,0,INDEX({vmat},A{r}-2005,MATCH({s},{vhd},0)))", EUR)
        output_cell(ws, r, 31, f"=IF(B{r}=0,0,Z{r}/AD{r}-1)", PCT2)

    section_header(ws, VM_SEC, "Prova: valore del titolo a fine anno nel motore mensile (dato storico del motore, non calcolato dal foglio)", span=12)
    intest(ws, VM_HDR, 1, ["Anno \\ partenza"] + PARTENZE, altezza=30)
    BT = B.dati_btp()
    for i, y in enumerate(ANNI):
        r = VM_Y0 + i
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        for k, p_ in enumerate(PARTENZE):
            d = BT.get((p_, y))
            if d is not None:
                input_cell(ws, r, 2 + k, float(d["valore_motore"]), EUR)
            else:
                ws.cell(row=r, column=2 + k).border = BORDER
    B.autofit(ws, max_w=30)
    ws.column_dimensions["A"].width = 46
    ws.freeze_panes = "B3"
