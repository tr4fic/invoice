"""Prevede vygenerovanou fakturu do PDF a posle ji emailem.

Pouziti:
    uv run send_invoice.py <mesic 1-12> [--rok ROK]

Faktura musi uz existovat v data/ (spust nejdriv create_invoice.py).
SMTP udaje se nastavuji v .env (viz .env.example).
"""

import argparse
import imaplib
import os
import re
import smtplib
import sys
from datetime import date
from email.message import EmailMessage
from pathlib import Path

import win32com.client as win32
from dotenv import load_dotenv

from create_invoice import MESICE, OUTPUT_DIR, last_day_of_previous_month

ROOT = Path(__file__).parent
ODESILATEL_ANCHOICE = "petr@anchoice.cz"
PRIJEMCE_ANCHOICE = "martin@anchoice.cz"
ODESILATEL_GMAIL = "erbenpet@gmail.com"
PRIJEMCE_GMAIL = "vnetik@deok.cz"
GMAIL_SMTP_HOST = "smtp.gmail.com"
GMAIL_SMTP_PORT = 587


def parse_args():
    parser = argparse.ArgumentParser(description="Prevede fakturu na PDF a odesle emailem.")
    parser.add_argument("mesic", type=int, choices=range(1, 13), help="cislo mesice 1-12")
    parser.add_argument("--rok", type=int, default=date.today().year, help="rok faktury (vychozi: aktualni rok)")
    return parser.parse_args()


def xls_to_pdf(xls_path: Path) -> Path:
    pdf_path = xls_path.with_suffix(".pdf")
    excel = win32.DispatchEx("Excel.Application")
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


def _attach_pdf(msg: EmailMessage, pdf_path: Path):
    msg.add_attachment(
        pdf_path.read_bytes(),
        maintype="application",
        subtype="pdf",
        filename=pdf_path.name,
    )


def build_email_anchoice(nazev_mesice: str, rok: int, pdf_path: Path) -> EmailMessage:
    rok_zkraceny = f"{rok % 100:02d}"
    msg = EmailMessage()
    msg["From"] = ODESILATEL_ANCHOICE
    msg["To"] = PRIJEMCE_ANCHOICE
    msg["Subject"] = f"Petr - faktura - {nazev_mesice} {rok_zkraceny}"
    msg.set_content(
        "Dobrý den, pane Krčmáři,\n\n"
        f"posílám fakturu za programátorské práce za měsíc {nazev_mesice} {rok}. Prosím o uhrazení.\n\n"
        "Díky a přeji hezký den,\n"
        "Petr Erben"
    )
    _attach_pdf(msg, pdf_path)
    return msg


def build_email_gmail(nazev_mesice: str, rok: int, pdf_path: Path) -> EmailMessage:
    rok_zkraceny = f"{rok % 100:02d}"
    msg = EmailMessage()
    msg["From"] = ODESILATEL_GMAIL
    msg["To"] = PRIJEMCE_GMAIL
    msg["Subject"] = f"Erbi - faktura {nazev_mesice} {rok_zkraceny}"
    msg.set_content(
        "Ahoj Vaso, posilam fakturu za AnChoice.\n"
        "Diky moc zpracovani.\n"
        "Erbi"
    )
    _attach_pdf(msg, pdf_path)
    return msg


def load_config() -> dict:
    load_dotenv(ROOT / ".env")
    cfg = {
        "smtp_host": os.environ.get("SMTP_HOST"),
        "smtp_port": os.environ.get("SMTP_PORT"),
        "user": os.environ.get("SMTP_USER"),
        "password": os.environ.get("SMTP_PASS"),
        "imap_host": os.environ.get("IMAP_HOST"),
        "imap_port": os.environ.get("IMAP_PORT"),
        "gmail_app_password": os.environ.get("GMAIL_APP_PASSWORD"),
    }
    if not all([cfg["smtp_host"], cfg["smtp_port"], cfg["user"], cfg["password"]]):
        sys.exit(
            "Chybi SMTP nastaveni. Vyplň .env v korenu repozitare podle .env.example:\n"
            "SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS"
        )
    if not cfg["gmail_app_password"]:
        sys.exit(
            "Chybi GMAIL_APP_PASSWORD v .env (app heslo pro erbenpet@gmail.com,\n"
            "viz .env.example) -- potreba pro odeslani druheho emailu."
        )
    return cfg


def send_email(msg: EmailMessage, host: str, port: int, user: str, password: str):
    smtp_class = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    with smtp_class(host, port) as smtp:
        if port != 465:
            smtp.starttls()
        smtp.login(user, password)
        smtp.send_message(msg)


def append_to_sent(msg: EmailMessage, cfg: dict):
    """Ulozi kopii odeslaneho emailu do slozky Odeslane, aby se objevila v Outlooku."""
    if not cfg["imap_host"] or not cfg["imap_port"]:
        return
    imap = imaplib.IMAP4_SSL(cfg["imap_host"], int(cfg["imap_port"]))
    try:
        imap.login(cfg["user"], cfg["password"])
        typ, folders = imap.list()
        sent_folder = "Sent"
        if typ == "OK":
            for line in folders:
                text = line.decode(errors="ignore")
                if "\\Sent" in text:
                    match = re.search(r'"\." "?([^"\r\n]+)"?$', text)
                    if match:
                        sent_folder = match.group(1)
                    break
        imap.append(sent_folder, "\\Seen", None, msg.as_bytes())
    finally:
        imap.logout()


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

    cfg = load_config()

    msg_anchoice = build_email_anchoice(nazev_mesice, rok_mesice, pdf_path)
    send_email(msg_anchoice, cfg["smtp_host"], int(cfg["smtp_port"]), cfg["user"], cfg["password"])
    print(f"Email odeslan na {PRIJEMCE_ANCHOICE}")
    append_to_sent(msg_anchoice, cfg)
    print("Kopie ulozena do Odeslane posty")

    msg_gmail = build_email_gmail(nazev_mesice, rok_mesice, pdf_path)
    send_email(msg_gmail, GMAIL_SMTP_HOST, GMAIL_SMTP_PORT, ODESILATEL_GMAIL, cfg["gmail_app_password"])
    print(f"Email odeslan na {PRIJEMCE_GMAIL}")


if __name__ == "__main__":
    main()
