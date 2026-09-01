"""Prevede vygenerovanou fakturu do PDF a posle ji emailem.

Pouziti:
    uv run send_invoice.py <mesic 1-12> [--rok ROK]

Faktura musi uz existovat v data/ (spust nejdriv create_invoice.py).
SMTP udaje se nastavuji v .env (viz .env.example).
"""

import argparse
import os
import smtplib
import sys
from datetime import date
from email.message import EmailMessage
from pathlib import Path

import win32com.client as win32
from dotenv import load_dotenv

from create_invoice import MESICE, OUTPUT_DIR, last_day_of_previous_month

ROOT = Path(__file__).parent
ODESILATEL = "petr@anchoice.cz"
PRIJEMCE = "erbenpet@gmail.com"


def parse_args():
    parser = argparse.ArgumentParser(description="Prevede fakturu na PDF a odesle emailem.")
    parser.add_argument("mesic", type=int, choices=range(1, 13), help="cislo mesice 1-12")
    parser.add_argument("--rok", type=int, default=date.today().year, help="rok faktury (vychozi: aktualni rok)")
    return parser.parse_args()


def xls_to_pdf(xls_path: Path) -> Path:
    pdf_path = xls_path.with_suffix(".pdf")
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(xls_path))
        ws = wb.Worksheets("faktura")
        ws.ExportAsFixedFormat(0, str(pdf_path))  # 0 = xlTypePDF
        wb.Close(False)
    finally:
        excel.Quit()
    return pdf_path


def build_email(nazev_mesice: str, rok: int, pdf_path: Path) -> EmailMessage:
    rok_zkraceny = f"{rok % 100:02d}"
    msg = EmailMessage()
    msg["From"] = ODESILATEL
    msg["To"] = PRIJEMCE
    msg["Subject"] = f"Petr - faktura - {nazev_mesice} {rok_zkraceny}"
    msg.set_content(
        "Dobrý den, pane Krčmáři,\n\n"
        f"posílám fakturu za programátorské práce za měsíc {nazev_mesice} {rok}. Prosím o uhrazení.\n\n"
        "Díky a přeji hezký den,\n"
        "Petr Erben"
    )
    msg.add_attachment(
        pdf_path.read_bytes(),
        maintype="application",
        subtype="pdf",
        filename=pdf_path.name,
    )
    return msg


def send_email(msg: EmailMessage):
    load_dotenv(ROOT / ".env")
    host = os.environ.get("SMTP_HOST")
    port = os.environ.get("SMTP_PORT")
    user = os.environ.get("SMTP_USER")
    password = os.environ.get("SMTP_PASS")
    if not all([host, port, user, password]):
        sys.exit(
            "Chybi SMTP nastaveni. Vyplň .env v korenu repozitare podle .env.example:\n"
            "SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS"
        )
    port = int(port)
    smtp_class = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    with smtp_class(host, port) as smtp:
        if port != 465:
            smtp.starttls()
        smtp.login(user, password)
        smtp.send_message(msg)


def main():
    args = parse_args()
    cislo = f"{args.rok}{args.mesic:03d}"
    xls_path = OUTPUT_DIR / f"anc_{cislo}.xls"
    if not xls_path.exists():
        sys.exit(f"Faktura {xls_path} neexistuje. Nejdriv spust: uv run create_invoice.py {args.mesic} --rok {args.rok}")

    duzp = last_day_of_previous_month(args.rok, args.mesic)
    nazev_mesice = MESICE[duzp.month]
    rok_mesice = duzp.year

    pdf_path = xls_to_pdf(xls_path)
    print(f"Vytvoreno PDF: {pdf_path}")

    msg = build_email(nazev_mesice, rok_mesice, pdf_path)
    send_email(msg)
    print(f"Email odeslan na {PRIJEMCE}")


if __name__ == "__main__":
    main()
