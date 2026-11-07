"""Foglio «4 - Casa»: appartamento medio italiano affittato; canone, costi, imposte, conto dei Buoni del Tesoro, vendita."""
from __future__ import annotations

from _va_base import *  # noqa: F401,F403
import _va_base as B

HDR = 22
R0 = 23
P0C, CATC = "$B$5", "$B$6"
RES = dict(valore=9, costi=10, conto=11, plus=12, netto=13)


def row(i):
    return R0 + i


def build(ws):
    testata(ws, 18)
    s = inp("start")
    section_header(ws, 3, "La casa comprata: prezzo e base della tassa sulla casa", span=6)
    label_cell(ws, 5, 1, "Prezzo della casa (capitale diviso uno più i costi di acquisto)")
    output_cell(ws, 5, 2, f"={inp('cap')}/(1+{inp('costi_acq')})", EUR)
    label_cell(ws, 6, 1, "Base della tassa sulla casa (valore catastale, fisso in euro)")
    output_cell(ws, 6, 2, f"={inp('catasto')}*{P0C}", EUR)
    note_cell(ws, 7, 1, "Appartamento medio italiano: prezzi dall'indice nazionale, canone da rendimento lordo medio, aliquote medie nazionali della tassa sulla casa.", span=6)

    section_header(ws, 8, "Se vendesse a ottobre 2026 (euro nominali, netto di tasse e costi)", span=6)
    last = row(NY - 1)
    lab = {"valore": "Valore della casa", "costi": "Costi di vendita", "conto": "Conto dei Buoni del Tesoro (affitti netti accumulati)",
           "plus": "Imposta sulla plusvalenza (solo se si vende entro cinque anni)", "netto": "Netto finale"}
    v, c, p = RES["valore"], RES["costi"], RES["plus"]
    fx = {"valore": f"=F{last}", "costi": f"={inp('costi_vend')}*B{v}", "conto": f"=R{last}",
          "plus": f"=IF((2026-{s}+1)*12-2<=60,{inp('aliq_cg_old')}*MAX(0,B{v}*(1-{inp('costi_vend')})-{inp('cap')}),0)",
          "netto": f"=B{v}-B{c}+B{RES['conto']}-B{p}"}
    for k in RES:
        label_cell(ws, RES[k], 1, lab[k])
        output_cell(ws, RES[k], 2, fx[k], EUR)
    note_cell(ws, 14, 1, "Oltre cinque anni la plusvalenza è esente (L. 266/2005 c. 496): con partenza 2006-2016 l'imposta è zero. "
              "Scarto dal motore mensile (partenza 2006): -0,52%.", span=6)
    note_cell(ws, 15, 1, "Fino al 2010 si tassa l'85% del canone con IRPEF e addizionale, più l'1% di registro; dal 2011 vale la tassa fissa sugli affitti.", span=6)

    section_header(ws, HDR - 1, "Motore del calcolo, non modificare: anno per anno", span=18)
    intest(ws, HDR, 1, ["Anno", "In corso (1 = sì)", "Frazione d'anno", "Valore a inizio anno", "Variazione dei prezzi", "Valore a fine anno", "Valore medio",
                        "Rendimento lordo del canone", "Indice degli affitti a fine anno", "Canone lordo incassato", "Manutenzione", "Aliquota tassa sulla casa",
                        "Tassa sulla casa", "Imposte sul canone", "Affitto netto dell'anno", "Conto prima del bollo", "Imposta di bollo", "Conto a fine anno"], altezza=60)
    for i, y in enumerate(ANNI):
        r = row(i)
        pr = r - 1
        ws.cell(row=r, column=1, value=y).font = FONT_LABEL
        irp = f"IF(A{r}=2006,{inp('irpef06')},{inp('irpef0710')})"
        F = {
            2: (f"=IF(A{r}>={s},1,0)", INT),
            3: (f"={dref('f', i)}", "0.0000"),
            4: (f"=IF(B{r}=0,0,IF(A{r}={s},{P0C},F{pr}))", EUR),
            5: (f"={dref('casa', i)}", PCT2),
            6: (f"=IF(B{r}=0,0,D{r}*(1+E{r}))", EUR),
            7: (f"=IF(B{r}=0,0,(D{r}+F{r})/2)", EUR),
            8: (f"=IF(B{r}=0,0,{inp('rend0')}+({inp('rend1')}-{inp('rend0')})*(A{r}-2006+0.5)/{inp('interp')})", PCT3),
            9: (f"=IF(B{r}=0,0,IF(A{r}={s},1,I{pr})*(1+{dref('hicp', i)}))", "0.0000"),
            10: (f"=IF(B{r}=0,0,IF({inp('modo')}=\"HICP\",{inp('rend0')}*{P0C}*(IF(A{r}={s},1,I{pr})+I{r})/2*(1-{inp('sfitto')})*C{r},"
                 f"H{r}*G{r}*(1-{inp('sfitto')})*C{r}))", EUR),
            11: (f"=IF(B{r}=0,0,{inp('manut')}*G{r}*C{r})", EUR),
            12: (f"={fref('ici', i)}", "0.000%"),
            13: (f"=IF(B{r}=0,0,L{r}*{CATC}*{fref('molt', i)}*C{r})", EUR),
            14: (f"=IF(B{r}=0,0,IF(A{r}<=2010,({irp}+{inp('extra')}+{inp('addiz')})*{inp('ded')}*J{r}+{inp('reg')}*J{r},{inp('cedolare')}*J{r}))", EUR),
            15: (f"=IF(B{r}=0,0,J{r}-K{r}-M{r}-N{r})", EUR),
            16: (f"=IF(B{r}=0,0,IF(A{r}={s},0,R{pr})*(1+{dref('bot', i)}*(1-{inp('aliq_ts')})*C{r})+O{r})", EUR),
            17: (f"=IF(B{r}=0,0,IF(P{r}>0,{B.bollo_expr(i, f'P{r}', f'C{r}')},0))", EUR),
            18: (f"=IF(B{r}=0,0,P{r}-Q{r})", EUR),
        }
        for col, (fx_, fm) in F.items():
            output_cell(ws, r, col, fx_, fm)
    B.autofit(ws, max_w=30)
    ws.column_dimensions["A"].width = 52
    ws.freeze_panes = "B3"
