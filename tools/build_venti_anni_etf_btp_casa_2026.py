"""Costruisce simulatori/venti_anni_etf_btp_casa_2026.xlsx (STAGING): tre fratelli, 100.000 euro nel 2006, venti anni dopo (caso inventato, numeri storici reali).

Specifica: _lab_adversarial_2026_08/venti_anni_etf_btp_casa/11_ANALISI_v2.md sez. 5 (indice + 5 fogli; anno di partenza e capitale come input; titolo di Stato con PREZZO();
riga Casa; storia fiscale anno per anno) e 90_VERDETTO.md. Standard: _VINCOLI_FIRE.md «Formattazione dell'Excel» (titolo riga 1, disclaimer riga 2, giallo = input,
verde = formule, azzurro = intestazioni, autofit, nessuna macro, nessun marchio). Moduli: _va_base, _va_s1 ... _va_s5.
Dopo la build: python tools/_recalc_venti_anni_etf_btp_casa_2026.py e python tools/_verifica_g6_venti_anni_etf_btp_casa_2026.py."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _va_base import *  # noqa: E402,F401,F403
import _va_base as B  # noqa: E402
import _va_s1, _va_s2, _va_s3, _va_s4, _va_s5  # noqa: E402


def main():
    info = [
        (S1, "Input (anno di partenza, capitale, costi), dati storici annuali, titolo comprato, tabella fiscale anno per anno"),
        (S2, "Fondo mondiale a distribuzione e ad accumulo, azioni italiane, Buoni del Tesoro: valore anno per anno e vendita"),
        (S3, "Un titolo di Stato decennale tenuto e rinnovato: valore con PREZZO(), cedole, conto dei Buoni del Tesoro, vendita"),
        (S4, "La casa da affittare: canone, costi, tassa sulla casa, imposte, conto dei Buoni del Tesoro, vendita"),
        (S5, "Riepilogo delle sei strade, scarto dichiarato dal motore mensile, cali storici e controlli"),
    ]
    sources = [
        "Indici azionari (mondo intero e Italia, con dividendi, in euro), 31/12/2005 - 30/10/2026; verifica del 06/11/2026 (E: serie ricostruite, fonte secondaria)",
        "Istat: indice FOI dei prezzi al consumo (P); indice nazionale dei prezzi delle case e indice europeo degli affitti (P/E)",
        "Banca d'Italia: asta del titolo di Stato di dicembre 2005 (titolo reale), rendimento dei Buoni del Tesoro a 12 mesi (P). I rinnovi sono un decennale tipo con cedola uguale al rendimento di allora, arrotondato al quarto di punto (E)",
        "Nomisma (2006) e media Italia abitazioni civili 2026 per il rendimento lordo degli affitti (E: fonte secondaria)",
        "D.Lgs. 461/1997 e DL 66/2014 (imposta sostitutiva 12,5% / 20% / 26%), DL 138/2011, D.Lgs. 239/1996 (titoli di Stato) (P)",
        "DPR 642/1972 (bollo 34,20 euro), DL 98/2011, DL 201/2011 art. 19, L. 147/2013 c. 581 (bollo in percentuale) (P)",
        "L. 228/2012 (imposta sulle transazioni), D.Lgs. 23/2011 art. 3 (tassa fissa sugli affitti), L. 266/2005 c. 496 (plusvalenza della casa) (P)",
        "CALCOLI MIEI (E): il motore annuale dei fogli 2-4 e le ipotesi dichiarate (costi, sfitto, manutenzione). Gradi: P = fonte primaria, E = stima o calcolo mio",
    ]
    build_cover(B.wb, TITOLO, [(S0, "Questo indice")] + info, sources)
    ws0 = B.wb[S0]
    ws0["A2"] = DISC
    ws0.row_dimensions[2].height = 44
    rr = 5 + 1 + len(info) + 1 + 6 + 1 + len(sources) + 1 + 2
    extra = [
        "Caso inventato e illustrativo: i tre fratelli non esistono, i numeri storici sì. Non è un consiglio e non è una previsione. Prima di decidere parla con professionisti abilitati.",
        "Come si usa: cambia l'anno di partenza (2006-2016) e il capitale nel foglio 1; leggi il riepilogo nel foglio 5. Le celle gialle sono input, le verdi formule.",
        "Stima annuale: il video usa un motore mensile. Questo foglio replica la meccanica con passo annuale: scarto dal motore da -0,5% a +2,0% (fondo a distribuzione +1,97%), dichiarato nel foglio 5.",
        "Ottobre 2026 è una stima: i dati del 2026 arrivano a fine ottobre (dieci mesi su dodici). L'inflazione di ottobre è una stima fino al dato definitivo.",
        "Il titolo di Stato è uno solo, non un paniere: il primo è reale, i rinnovi sono un decennale tipo con cedola uguale al rendimento di allora, arrotondato al quarto di punto. Il suo valore viene da PREZZO() applicato a un rendimento di fine anno stimato (foglio 1, tabella C), non dal motore.",
        "La lettura (b), redditi incassati e spesi, non è nel foglio ma solo nel video. Per le azioni italiane si usa l'indice con dividendi; il passo annuale reinveste i dividendi a metà anno.",
        "Non incluse: TASI 2014-2019 (effetto stimato ≤0,04 punti l'anno), spread denaro-lettera, successione, morosità, spese straordinarie.",
        "Bollo sul conto titoli: 34,20 € fino al 2010; 2011: 34,20 € fino al 16/07, poi 70 € sullo scaglione 50-150 mila € (DL 98/2011), ripartiti per giorni; dal 2012 uno per mille (DL 201/2011), 2 per mille dal 2014.",
        "Nessuna macro, solo formule. Le celle verdi non vanno modificate.",
    ]
    ws0.cell(row=rr - 1, column=1, value="Note e limiti").font = FONT_LABEL
    for i, t in enumerate(extra):
        ws0.cell(row=rr + i, column=1, value=t).font = FONT_VALUE
    ws1, ws2, ws3, ws4, ws5 = (B.wb.create_sheet(nm) for nm in (S1, S2, S3, S4, S5))
    _va_s1.build(ws1)
    _va_s2.build(ws2)
    _va_s3.build(ws3)
    _va_s4.build(ws4)
    _va_s5.build(ws5)
    B.autofit(ws0, max_w=110)
    ws0.column_dimensions["A"].width = 112
    OUT.parent.mkdir(parents=True, exist_ok=True)
    B.wb.save(OUT)
    print("scritto", OUT, OUT.stat().st_size, "byte")


if __name__ == "__main__":
    main()
