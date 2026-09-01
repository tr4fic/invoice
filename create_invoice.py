"""Vygeneruje novou fakturu ze sablony pro dany mesic.

Pouziti:
    uv run create_invoice.py <mesic 1-12> [--rok ROK]

Priklad:
    uv run create_invoice.py 1
    -> vytvori data/anc_2026001.xls (faktura za prosinec 2025, splatnost 15.1.2026)
"""

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

import win32com.client as win32

ROOT = Path(__file__).parent
TEMPLATE = ROOT / "templates" / "anc_sablona.xls"
OUTPUT_DIR = ROOT / "data"

MESICE = {
    1: "leden", 2: "únor", 3: "březen", 4: "duben", 5: "květen", 6: "červen",
    7: "červenec", 8: "srpen", 9: "září", 10: "říjen", 11: "listopad", 12: "prosinec",
}

CASTKA_PROGRAMATOR = 132000
CASTKA_CLAUDE_PRO = 550


def parse_args():
    parser = argparse.ArgumentParser(description="Vytvori novou fakturu ze sablony.")
    parser.add_argument("mesic", type=int, choices=range(1, 13), help="cislo mesice 1-12")
    parser.add_argument("--rok", type=int, default=date.today().year, help="rok faktury (vychozi: aktualni rok)")
    return parser.parse_args()


def last_day_of_previous_month(year: int, month: int) -> date:
    return date(year, month, 1) - timedelta(days=1)


EXCEL_EPOCH = date(1899, 12, 30)


def excel_serial(d: date) -> int:
    return (d - EXCEL_EPOCH).days


def build_invoice(mesic: int, rok: int):
    cislo = f"{rok}{mesic:03d}"

    datum_vystaveni = date.today()
    duzp = last_day_of_previous_month(rok, mesic)
    splatnost = date(rok, mesic, 15)

    nazev_mesice = MESICE[duzp.month]
    rok_mesice = duzp.year

    if not TEMPLATE.exists():
        sys.exit(f"Sablona nenalezena: {TEMPLATE}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    output_path = OUTPUT_DIR / f"anc_{cislo}.xls"

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(TEMPLATE))
        ws = wb.Worksheets("faktura")

        ws.Range("I1").Value = f"číslo: {cislo}"
        ws.Range("I16").Value = excel_serial(datum_vystaveni)
        ws.Range("I17").Value = excel_serial(duzp)
        ws.Range("I18").Value = excel_serial(splatnost)
        ws.Range("E19").Value = int(cislo)
        ws.Range("C23").Value = f"Programátorské práce na programu Anchoice za {nazev_mesice} {rok_mesice}."
        ws.Range("I23").Value = CASTKA_PROGRAMATOR
        ws.Range("C24").Value = "Claude Pro"
        ws.Range("I24").Value = CASTKA_CLAUDE_PRO

        excel.CalculateFull()

        wb.SaveAs(str(output_path), FileFormat=56)  # xlExcel8 (.xls)
        wb.Close(False)
    finally:
        excel.Quit()

    return output_path, cislo


def main():
    args = parse_args()
    output_path, cislo = build_invoice(args.mesic, args.rok)
    print(f"Vytvorena faktura {cislo}: {output_path}")


if __name__ == "__main__":
    main()
