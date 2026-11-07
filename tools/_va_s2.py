"""Foglio «2 - Borsa»: fondo a distribuzione, fondo ad accumulo, azioni italiane, Buoni del Tesoro. Risultati in alto, motore anno per anno in fondo."""
from __future__ import annotations

from _va_base import *  # noqa: F401,F403
import _va_base as B

RES_HDR = 4
RES_ROWS = {"dist": 5, "acc": 6, "it": 7, "bot": 8}
V0_ROW = 10
V0 = f"$B${V0_ROW}"
BLK = {}        # nome blocco -> riga della cella 2005 (r0)
TOP = {"dist": 15, "acc": 15 + 26, "it": 15 + 52, "bot": 15 + 78}


def r_of(name, i):
    return BLK[name] + 1 + i


def _act(r):
    return f"$A{r}>={inp('start')}"


def _fine(r, calc, cond_extra=None):
    """Valore di fine anno: 0 prima di (partenza-1), capitale iniziale in (partenza-1), poi il calcolo."""
    return f"=IF($A{r}<{inp('start')}-1,0,IF($A{r}={inp('start')}-1,{V0},{calc}))"


def _blocco_azioni(ws, name, titolo, kc, kt, ter, tobin):
    top = TOP[name]
    BLK[name] = top + 2
    section_header(ws, top, titolo, span=14)
    cols = ["Anno", "Frazione d'anno", "Rendimento del prezzo", "Rendimento con dividendi", "Valore a inizio anno", "Valore dopo prezzo e costo del fondo",
            "Dividendi lordi", "Dividendi dopo l'imposta"]
    if tobin:
        cols.append("Imposta sulle transazioni")
    cols += ["Dividendi reinvestiti, dopo commissione", "Valore prima del bollo", "Imposta di bollo", "Valore a fine anno", "Costo fiscale a fine anno"]
    intest(ws, top + 1, 1, cols, altezza=60)
    # lettere colonne
    nm = ["A", "B", "C", "D", "E", "F", "G", "H"]
    if tobin:
        nm.append("I")
    base = len(nm)
    cJ, cK, cL, cM, cN = (L(base + 1 + k) for k in range(5))
    cI = "I" if tobin else None
    r0 = BLK[name]
    ws.cell(row=r0, column=1, value=2005).font = FONT_LABEL
    output_cell(ws, r0, base + 4, _fine(r0, "0"), EUR)
    output_cell(ws, r0, base + 5, _fine(r0, "0"), EUR)
    ter_f = f"(1-{inp('ter')}*B{{r}})" if ter else "1"
    for i, y in enumerate(ANNI):
        r = r0 + 1 + i
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        a = _act(r)
        output_cell(ws, r, 2, f"={dref('f', i)}", "0.0000")
        output_cell(ws, r, 3, f"={dref(kt, i)}", PCT2)
        output_cell(ws, r, 4, f"={dref(kc, i)}", PCT2)
        output_cell(ws, r, 5, f"={cM}{r-1}", EUR)
        tf = ter_f.format(r=r)
        output_cell(ws, r, 6, f"=IF({a},E{r}*(1+C{r})*{tf},0)", EUR)
        output_cell(ws, r, 7, f"=IF({a},E{r}*(D{r}-C{r})*{tf},0)", EUR)
        output_cell(ws, r, 8, f"=G{r}*(1-{fref('aliq', i)})", EUR)
        if tobin:
            output_cell(ws, r, 9, f"=H{r}*{fref('tobin', i)}", EUR)
            sott = f"-{cI}{r}"
        else:
            sott = ""
        output_cell(ws, r, base + 1, f"=IF({a},H{r}{sott}-{B.fee_expr(f'H{r}')},0)", EUR)
        output_cell(ws, r, base + 2, f"=IF({a},F{r}+{cJ}{r}*(1+C{r}/2*B{r}),0)", EUR)
        output_cell(ws, r, base + 3, f"=IF({a},{B.bollo_expr(i, f'{cK}{r}', f'B{r}')},0)", EUR)
        output_cell(ws, r, base + 4, _fine(r, f"{cK}{r}-{cL}{r}"), EUR)
        output_cell(ws, r, base + 5, _fine(r, f"{cN}{r-1}+H{r}{sott}"), EUR)
    BLK[name + "_cols"] = (cM, cN)


def _blocco_acc(ws):
    top = TOP["acc"]
    BLK["acc"] = top + 2
    section_header(ws, top, "Fondo indicizzato sul mondo intero, ad accumulo (variante: i dividendi restano dentro)", span=7)
    intest(ws, top + 1, 1, ["Anno", "Frazione d'anno", "Rendimento con dividendi", "Valore a inizio anno", "Valore prima del bollo", "Imposta di bollo",
                            "Valore a fine anno"], altezza=60)
    r0 = BLK["acc"]
    ws.cell(row=r0, column=1, value=2005).font = FONT_LABEL
    output_cell(ws, r0, 7, _fine(r0, "0"), EUR)
    for i, y in enumerate(ANNI):
        r = r0 + 1 + i
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        output_cell(ws, r, 2, f"={dref('f', i)}", "0.0000")
        output_cell(ws, r, 3, f"={dref('w_net', i)}", PCT2)
        output_cell(ws, r, 4, f"=G{r-1}", EUR)
        output_cell(ws, r, 5, f"=IF({_act(r)},D{r}*(1+C{r})*(1-{inp('ter')}*B{r}),0)", EUR)
        output_cell(ws, r, 6, f"=IF({_act(r)},{B.bollo_expr(i, f'E{r}', f'B{r}')},0)", EUR)
        output_cell(ws, r, 7, _fine(r, f"E{r}-F{r}"), EUR)


def _blocco_bot(ws):
    top = TOP["bot"]
    BLK["bot"] = top + 2
    section_header(ws, top, "Buoni del Tesoro a 12 mesi (rinnovati ogni anno)", span=7)
    intest(ws, top + 1, 1, ["Anno", "Frazione d'anno", "Rendimento lordo", "Valore a inizio anno", "Valore prima del bollo", "Imposta di bollo",
                            "Valore a fine anno"], altezza=60)
    r0 = BLK["bot"]
    ws.cell(row=r0, column=1, value=2005).font = FONT_LABEL
    output_cell(ws, r0, 7, _fine(r0, "0"), EUR)
    for i, y in enumerate(ANNI):
        r = r0 + 1 + i
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        output_cell(ws, r, 2, f"={dref('f', i)}", "0.0000")
        output_cell(ws, r, 3, f"={dref('bot', i)}", PCT2)
        output_cell(ws, r, 4, f"=G{r-1}", EUR)
        output_cell(ws, r, 5, f"=IF({_act(r)},D{r}*(1+C{r}*(1-{inp('aliq_ts')})*B{r}),0)", EUR)
        output_cell(ws, r, 6, f"=IF({_act(r)},{B.bollo_expr(i, f'E{r}', f'B{r}')},0)", EUR)
        output_cell(ws, r, 7, _fine(r, f"E{r}-F{r}"), EUR)


def build(ws):
    testata(ws, 14)
    # blocchi del motore (in fondo)
    _blocco_azioni(ws, "dist", "Motore del calcolo, non modificare: fondo indicizzato sul mondo intero, a distribuzione", "w_net", "w_prc", True, False)
    _blocco_acc(ws)
    _blocco_azioni(ws, "it", "Motore del calcolo, non modificare: azioni italiane (indice con dividendi, senza costo del fondo)", "i_gr", "i_prc", False, True)
    _blocco_bot(ws)
    # risultati in alto
    section_header(ws, 3, "Risultati: se vendesse a ottobre 2026 (euro nominali, netto di tasse e costi)", span=8)
    intest(ws, RES_HDR, 1, ["Strada", "Valore lordo", "Costo fiscale", "Guadagno imponibile", "Commissione di vendita", "Imposta sul guadagno", "Netto finale",
                            "Guadagno maturato al 30/06/2014"], altezza=45)
    label_cell(ws, V0_ROW, 1, "Valore iniziale dopo la commissione di acquisto")
    output_cell(ws, V0_ROW, 2, f"={inp('cap')}-{B.fee_expr(inp('cap'))}", EUR)
    note_cell(ws, V0_ROW, 3, "Capitale meno la commissione sull'acquisto: parte da qui ogni strada della Borsa.", span=6)
    i13, i14 = ANNI.index(2013), ANNI.index(2014)
    righe = (("dist", "Fondo mondiale, a distribuzione"), ("acc", "Fondo mondiale, ad accumulo"), ("it", "Azioni italiane"))
    for name, nome in righe:
        r = RES_ROWS[name]
        label_cell(ws, r, 1, nome)
        if name == "acc":
            v = f"=G{r_of('acc', NY - 1)}"
            cst = f"={V0}"
            m13, m14 = f"G{r_of('acc', i13)}", f"G{r_of('acc', i14)}"
            p14 = f"=MAX(0,0.5*(IF({m13}>0,{m13}-{V0},0)+IF({m14}>0,{m14}-{V0},0)))"
        else:
            cM, cN = BLK[name + "_cols"]
            v = f"={cM}{r_of(name, NY - 1)}"
            cst = f"={cN}{r_of(name, NY - 1)}"
            p14 = (f"=MAX(0,0.5*(({cM}{r_of(name, i13)}-{cN}{r_of(name, i13)})+({cM}{r_of(name, i14)}-{cN}{r_of(name, i14)})))")
        output_cell(ws, r, 2, v, EUR)
        output_cell(ws, r, 3, cst, EUR)
        output_cell(ws, r, 4, f"=B{r}-C{r}", EUR)
        output_cell(ws, r, 5, f"={B.fee_expr(f'B{r}')}", EUR)
        output_cell(ws, r, 8, p14, EUR)
        output_cell(ws, r, 6, f"=IF(D{r}>0,{inp('aliq_cg')}*D{r}-({inp('aliq_cg')}-{inp('aliq_cg_old')})*MIN(D{r},H{r}),0)", EUR)
        output_cell(ws, r, 7, f"=B{r}-E{r}-F{r}", EUR)
    r = RES_ROWS["bot"]
    label_cell(ws, r, 1, "Buoni del Tesoro a 12 mesi")
    output_cell(ws, r, 2, f"=G{r_of('bot', NY - 1)}", EUR)
    for c in (3, 4, 5, 6, 8):
        cc = ws.cell(row=r, column=c, value="non si applica")
        cc.font = FONT_NOTE
        cc.border = BORDER
    output_cell(ws, r, 7, f"=B{r}", EUR)
    note_cell(ws, 9, 1, "Imposta sui guadagni: aliquota piena sul guadagno, meno lo sconto per la quota maturata al 30/06/2014 (era tassata al 20%). Il fondo a distribuzione paga già "
              "l'imposta sui dividendi anno per anno. Ad accumulo la paga tutta alla vendita.", span=8)
    note_cell(ws, 12, 1, "Scarto del foglio annuale dal motore mensile del video (partenza 2006): fondo a distribuzione +1,97%, accumulo +0,75%, azioni italiane +0,71%, "
              "Buoni del Tesoro -0,32%. Dichiarato nel foglio 5.", span=8)
    B.autofit(ws, max_w=26)
    ws.column_dimensions["A"].width = 46
    ws.freeze_panes = "B3"
