#!/usr/bin/env python3
"""Convert Comdirect Finanzreport PDFs in this script's folder to one UTF-8 CSV.

Dependency: pypdf (install with: python -m pip install pypdf)
Run: python parse_finanzreport.py
"""

import csv
import re
from bisect import bisect_right
from datetime import datetime
from pathlib import Path
from statistics import median

from pypdf import PdfReader

ACCOUNT_TITLES = ("Girokonto", "Visa-Karte")
TABLE_HEADERS = ("Buchungstag", "Vorgang", "Auftraggeber", "Buchungstext", "Ausgang")
IGNORED_SUBHEADERS = ("Valuta", "Umsatztag")
PRINT_HEADERS = (
    "Buchungstag (orig)",
    "Vorgang (orig)",
    "Auftraggeber (orig)",
    "Buchungstext (orig)",
    "Ausgang/Eingang (orig)",
)

END_WORD = "(?:[^a-zA-Z]|$)"
CATEGORY_REGEX = [
    ("Vorgang (orig)", r"(?i:wertpapiere)", "Depot", "Kauf / Verkauf"),
    ("Vorgang (orig)", r"(?i:Termingeld)", "Depot", "Kauf / Verkauf"),
    ("Vorgang (orig)", r"(?:Bar)", "Einkommen", "Bargeldeinzahlung"),
    ("Vorgang (orig)", r"(?i:auszahlung gaa)|(?i:barauszahlung)", "Sonstige Ausgaben", "Bargeldauszahlung"),
    ("Auftraggeber-Name", r"comdirect Visa", "Ausschließen", "Interner Übertrag"),
    ("Auftraggeber-Name", r"(?i:stadtwerke)|(?i:abfall)", "Wohnen", "Nebenkosten"),
    ("Auftraggeber-Name", r"(?i:kaufland)|(?i:lidl)|(?i:aldi)|(?i:rewe)|(?i:penny)|(?i:edeka)|(?i:tegut)", "Lebenshaltung", "Lebensmittel"),
    ("Auftraggeber-Name", r"(?i:drogeriemarkt)|(?i:rossmann)", "Lebenshaltung", "Drogerie"),
    ("Auftraggeber-Name", r"(?i:deichmann)|(?i:c\+a)", "Lebenshaltung", "Kleidung / Schuhe"),
    ("Auftraggeber-Name", rf"(?i:obi{END_WORD})|(?i:bauhaus{END_WORD})", "Lebenshaltung", "Heim / Garten"),
    ("Auftraggeber-Name", r"(?i:apotheke)", "Lebenshaltung", "Medikamente"),
    ("Auftraggeber-Name", r"(?i:db vertrieb)|(?i:deutsche bahn)|(?i:rnv)", "Verkehrsmittel", "Öffentliche Verkehrsmittel"),
    ("Auftraggeber-Name", r"(?:OIL)", "Verkehrsmittel", "Auto / Tanken"),
    ("Auftraggeber-Name", r"(?i:unitymedia)|(?i:vodafone)|(?i:telefonica)|(?i:drillisch)|(?i:congstar)", "Digital", "Internet / Telefon"),
    ("Auftraggeber-Name", r"(?i:rundfunk)", "Digital", "Rundfunksteuer"),
    ("Auftraggeber-Name", rf"(?i:mcdonalds)|(?i:kfc{END_WORD})|(?i:nordsee gmbh)|(?i:moschmosch)|(?i:gastronomie)|(?i:restaurant)", "Freizeit", "Gastronomie"),
    ("Auftraggeber-Name", r"(?i:cineplex)|(?i:filmpalast)", "Freizeit", "Unterhaltung / Kino / Kultur"),
    ("Auftraggeber-Name", r"(?i:germanwings)", "Reisen", "Flug / Bahn / Bus / Taxi"),
    ("Auftraggeber-Name", r"(?i:bundesagentur für arbeit)", "Einkommen", "Arbeitslosengeld"),
    ("Buchungsnotiz", r"(?i:darlehen)", "Darlehen", "Tilgung"),
    ("Buchungsnotiz", r"(?i:miete)", "Wohnen", "Miete"),
    ("Buchungsnotiz", r"(?i:nebenkostenabrechnung)", "Wohnen", "Miete"),
    ("Buchungsnotiz", r"(?i:hausgeld)", "Wohnen", "Hausgeld"),
    ("Buchungsnotiz", r"(?i:strom)|(?i:wasser)|(?i:heizung)|(?i:fernwaerme)|(?i:abfall)", "Wohnen", "Nebenkosten"),
    ("Buchungsnotiz", rf"(?i:rewe{END_WORD})", "Lebenshaltung", "Lebensmittel"),
    ("Buchungsnotiz", rf"(?i:zara{END_WORD})|(?i:primark)|(?i:Mountain Sports Outlet GmbH)|(?i:Engelhor\s?n Mode GmbH)", "Lebenshaltung", "Kleidung / Schuhe"),
    ("Buchungsnotiz", r"(?i:uebertrag auf)|(?i:tagesgeld)|(?i:ueberweisung auf v\s?isa-kar\s?te)|(?i:summe wochenabrechnung \s?visa)|(?i:summe monatsabrechnung visa)|(?i:Ueberweisung von Girokonto auf Visa-Karte)", "Ausschließen", "Interner Übertrag"),
    ("Buchungsnotiz", rf"(?i:wage{END_WORD})|(?i:salary)|(?i:gehalt)|(?i:lohn)|(?i:bezuege )", "Einkommen", "Gehalt"),
    ("Buchungsnotiz", r"(?i:ertraegnisgutschrift)", "Einkommen", "Sparen / Anlegen"),
    ("Buchungsnotiz", r"(?i:bargeldeinzahlung)", "Einkommen", "Bargeldeinzahlung"),
    ("Buchungsnotiz", r"(?i:bargeldauszahlung)", "Sonstige Ausgaben", "Bargeldauszahlung"),
    ("Buchungsnotiz", r"(?i:apotheke)", "Lebenshaltung", "Medikamente"),
    ("Buchungsnotiz", r"(?i:dbvertrieb)", "Verkehrsmittel", "Öffentliche Verkehrsmittel"), # when paying via Paypal
    ("Buchungsnotiz", r"(?i:ryanair)|(?i:wizz air)|(?i:tarom)|(?i:fluegede)|(?i:lufthansa)|(?i:Austrian Airlines AG)|(?i:Luchtvaart Maatschappij)", "Reisen", "Flug / Bahn / Bus / Taxi"), # when paying via Paypal
    ("Buchungsnotiz", r"(?i:hotel)", "Reisen", "Hotel / Unterkunft"), # when paying by VISA
    ("Buchungsnotiz", r"(?i:burger king)|(?i:rieger gastro)", "Freizeit", "Gastronomie"),
    ("Buchungsnotiz", r"(?i:theater)|(?i:eventim)|(?i:fkp scorpio)", "Freizeit", "Unterhaltung / Kino / Kultur"), # when paying by VISA
    ("Buchungsnotiz", r"(?i:humblebundl)|(?i:steam\s?games)|(?i:stea\s?mpowered)|(?i:wargaming)", "Freizeit", "Hobbies"), # when paying via Paypal
]


# Lightweight replacement for the pandas DataFrame used by the original project.
class AccountTable:
    def __init__(self):
        self.columns = []
        self.rows = {}

    @property
    def empty(self):
        return not self.rows or not self.columns

    def set(self, row, column, value):
        if column not in self.columns:
            self.columns.append(column)
        self.rows.setdefault(row, {})[column] = value

    def get(self, row, column):
        return self.rows.get(row, {}).get(column)

    def ensure_column(self, column):
        if column not in self.columns:
            self.columns.append(column)

    def insert_blank_column(self, column):
        self.ensure_column(column)
        for row in self.rows.values():
            row.setdefault(column, "")


def _decode_pdf_text(value):
    if isinstance(value, str):
        return value
    try:
        return value.decode("cp1252")
    except (AttributeError, UnicodeDecodeError):
        try:
            return bytes(value).decode("cp1252", errors="replace")
        except Exception:
            return str(value)


def parse_finanzreport(pdf_path):
    ACCOUNT_END = "Neuer Saldo"
    ACCOUNT_END_SIGNAL_1 = "Neuer"
    ACCOUNT_END_SIGNAL_2 = "Saldo"
    LAST_HEADER_LEEWAY = 10
    TITLE_FONT_SIZE_THRESHOLD = 9
    ANY_FONT_SIZE_THRESHOLD = 3
    LINE_BREAK_THRESHOLD = 12.01

    account_tables = {title: AccountTable() for title in ACCOUNT_TITLES}
    active_account = ""
    header_xcoords = {title: [] for title in ACCOUNT_TITLES}
    seen_header_texts = {title: [] for title in ACCOUNT_TITLES}
    end_signal_1_seen = False
    old_pdf_font_size = 0
    first_header = 0
    previous_column = 0.0
    current_row_y = 10_000_000.0
    first_column_line_break_y = 0.0
    current_row = 0
    row_does_not_start_with_date = True

    def find_closest_header(text_x, headers):
        index = bisect_right(headers, text_x)
        return None if index == 0 else headers[index - 1]

    def is_date(text):
        return re.match(r"A?\d{2}\.\d{2}\.\d{4}", text) is not None

    def is_account_title(text, font_size):
        return font_size > TITLE_FONT_SIZE_THRESHOLD and any(t in text for t in ACCOUNT_TITLES)

    def interpret_chunk(operator, operands, _transformation_matrix, text_matrix):
        nonlocal old_pdf_font_size, active_account, end_signal_1_seen
        nonlocal first_header, previous_column, current_row, current_row_y
        nonlocal first_column_line_break_y, row_does_not_start_with_date

        if operator not in (b"TJ", b"Tj"):
            if operator == b"Tf" and len(operands) > 1:
                old_pdf_font_size = operands[1]
            return

        try:
            raw_text = operands[0][0] if operator == b"TJ" else operands[0]
            text = _decode_pdf_text(raw_text)
        except (IndexError, TypeError):
            return

        font_size = text_matrix[0]
        if font_size < ANY_FONT_SIZE_THRESHOLD and old_pdf_font_size < ANY_FONT_SIZE_THRESHOLD:
            return
        if font_size < ANY_FONT_SIZE_THRESHOLD:
            font_size = old_pdf_font_size

        x, y = text_matrix[4], text_matrix[5]
        if is_account_title(text, font_size):
            active_account = next(t for t in ACCOUNT_TITLES if t in text)
            first_header = 0
            previous_column = 0.0
            current_row_y = 10_000_000.0
            current_row = 0
            row_does_not_start_with_date = True
            return

        if not active_account:
            return

        table = account_tables[active_account]
        headers = header_xcoords[active_account]

        if any(header in text for header in TABLE_HEADERS):
            is_last_header = TABLE_HEADERS[-1] in text
            header_x = x - LAST_HEADER_LEEWAY if is_last_header else x
            if header_x not in headers and text not in seen_header_texts[active_account]:
                headers.append(header_x)
                headers.sort()
                seen_header_texts[active_account].append(text)
            if TABLE_HEADERS[0] in text:
                first_header = x
                current_row_y = y
            return

        column = find_closest_header(x + LAST_HEADER_LEEWAY, headers)
        if column is None or y > current_row_y:
            return

        if text == ACCOUNT_END_SIGNAL_1:
            end_signal_1_seen = True
            return
        if text == ACCOUNT_END or (end_signal_1_seen and text == ACCOUNT_END_SIGNAL_2):
            active_account = ""
            return
        if end_signal_1_seen:
            end_signal_1_seen = False

        if column == first_header:
            if (current_row_y - y) < LINE_BREAK_THRESHOLD and current_row:
                first_column_line_break_y = y
                old_date = table.get(current_row, column)
                if text not in IGNORED_SUBHEADERS and text != old_date:
                    print(f"Info: Wertstellung eines Vorgangs am {text} != Buchungstag {old_date}.")
                return
            if not is_date(text):
                row_does_not_start_with_date = True
                return
            if text[0].isalpha():
                text = text[1:]
            current_row += 1
            current_row_y = first_column_line_break_y = y
            row_does_not_start_with_date = False

        if row_does_not_start_with_date:
            return

        old_value = table.get(current_row, column)
        if column == previous_column and old_value is not None:
            table.set(current_row, column, f"{old_value} {text}")
        else:
            table.set(current_row, column, text)
        previous_column = column

    reader = PdfReader(str(pdf_path), strict=False)
    for page in reader.pages:
        # visitor_operand_before is needed because ordinary extract_text loses spacing
        # information in these Finanzreport PDFs.
        page.extract_text(visitor_operand_before=interpret_chunk)

    for title, table in account_tables.items():
        headers = header_xcoords[title]
        if title == "Visa-Karte" and len(table.columns) == 4 and headers:
            table.insert_blank_column(median(headers))
        for x in headers:
            table.ensure_column(x)
        print(f"Extrahiert: {sum(1 for row in table.rows.values() if any(v not in (None, '') for v in row.values()))} Umsätze ({title}) aus {pdf_path.name}.")

    return account_tables


def parse_error(tables, filename):
    if all(table.empty for table in tables.values()):
        return f"Keine Umsätze gefunden in {filename}. Das PDF-Format könnte abweichen."
    for account, table in tables.items():
        if not table.empty and len(table.columns) != len(TABLE_HEADERS):
            return (f"Unerwartete Spaltenzahl für {account} in {filename}: "
                    f"{len(table.columns)} statt {len(TABLE_HEADERS)}. PDF wurde übersprungen.")
    return ""


def prettify_rows(table, filename, account_title):
    regex_ibanbic = re.compile(r"(?P<sender>.+?) (?P<ibanbic>[A-Z]{2}\d{2}[A-Z0-9]+ [A-Z]{6}[A-Z0-9]{5})")
    regex_reftext = re.compile(r"(?P<manualref>.*?) ?End-to-End-Ref\.:(?P<endref>.*)")
    sorted_columns = sorted(table.columns)
    if len(sorted_columns) < len(PRINT_HEADERS):
        raise ValueError(f"Zu wenige Tabellenspalten in {filename} ({account_title}).")

    output = []
    for row_index in sorted(table.rows):
        row_data = table.rows[row_index]
        raw = [row_data.get(column, "") or "" for column in sorted_columns[:5]]
        # Ignore empty placeholder rows.
        if not any(str(value).strip() for value in raw):
            continue
        record = dict(zip(PRINT_HEADERS, raw))
        record["Dateiname"] = filename
        record["Konto"] = account_title

        sender_text = str(record[PRINT_HEADERS[2]] or "")
        match = regex_ibanbic.search(sender_text)
        record["Auftraggeber-Name"] = match.group("sender") if match else sender_text
        record["Auftraggeber-IBAN/BIC"] = match.group("ibanbic") if match else ""

        booking_text = str(record[PRINT_HEADERS[3]] or "")
        match = regex_reftext.search(booking_text)
        record["Buchungsnotiz"] = match.group("manualref") if match else booking_text
        record["Buchungsreferenz"] = match.group("endref") if match else ""

        amount_text = str(record[PRINT_HEADERS[4]] or "").strip()
        normalized = amount_text.replace(".", "").replace(",", ".")
        try:
            record["Betrag (Zahl)"] = str(float(normalized)) if normalized else ""
        except ValueError:
            record["Betrag (Zahl)"] = ""

        record["Kategorie"] = ""
        record["Unterkategorie"] = ""
        for match_column, regex, category, subcategory in CATEGORY_REGEX:
            if re.search(regex, str(record.get(match_column, "") or "")):
                record["Kategorie"] = category
                record["Unterkategorie"] = subcategory
        output.append(record)
    return output

OUTPUT_HEADERS = list(PRINT_HEADERS) + [
    "Dateiname", "Konto", "Auftraggeber-Name", "Auftraggeber-IBAN/BIC",
    "Buchungsnotiz", "Buchungsreferenz", "Betrag (Zahl)", "Kategorie", "Unterkategorie",
]


SIMPLE_HEADERS = [
    "Konto",
    "Buchungstag",
    "Umsatzart",
    "Begünstigter / Auftraggeber",
    "Verwendungszweck",
    "Betrag",
    "Währung",
    "Sonstiges",
]

# Match the longest known transaction-type prefixes first.
TRANSACTION_TYPES = sorted(
    [
        "Lastschrift / Belastung",
        "Übertrag / Überweisung",
        "Gutschrift / Überweisung",
        "Überweisung / Gutschrift",
        "Wertpapierabrechnung",
        "Kartenverfügung",
        "Auszahlung GAA",
        "Kartenzahlung",
        "Devisen",
        "Dauerauftrag",
        "Wertpapiere",
        "Überweisung",
        "Gutschrift",
        "Lastschrift",
        "Entgelte",
        "Belastung",
        "Übertrag",
        "Auszahlung",
        "Einzahlung",
        "Zinsen",
        "Scheck",
    ],
    key=len,
    reverse=True,
)


def iso_date(value):
    """Convert German dates to ISO format, supporting two-digit years."""
    value = str(value or "").strip()
    match = re.fullmatch(r"A?(\d{2})\.(\d{2})\.(\d{2}|\d{4})", value)

    if not match:
        return value

    day, month, year = match.groups()
    if len(year) == 2:
        year = ("20" if int(year) <= 69 else "19") + year

    try:
        return datetime(
            int(year), int(month), int(day)
        ).strftime("%Y-%m-%d")
    except ValueError:
        return value



def extract_transaction_type(value):
    """Extract transaction type and keep the remaining text separately."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()

    # Repair the specific word split caused by PDF extraction.
    text = re.sub(
        r"(?i)\bkartenverfügun\s+g\b",
        "Kartenverfügung",
        text,
    )

    for transaction_type in TRANSACTION_TYPES:
        if text.casefold().startswith(transaction_type.casefold()):
            remainder = text[len(transaction_type):].strip()
            return transaction_type, remainder

    return text, ""

def make_simple_record(record):
    transaction_type, extra = extract_transaction_type(
        record.get("Vorgang (orig)", "")
    )

    booking_text = str(
        record.get("Buchungstext (orig)", "") or ""
    ).strip()

    beneficiary = str(
        record.get("Auftraggeber-Name", "") or ""
    ).strip()

    purpose = booking_text
    other_parts = [extra] if extra else []

    if transaction_type.casefold() == "kartenverfügung":
        transaction_type = "Kartenverfügung"

        # Separate the card details from the merchant description.
        marker = re.search(r"(?i)\bKarte\s+Nr\.", purpose)

        if marker:
            purpose = purpose[:marker.start()].strip()
            card_details = booking_text[marker.start():].strip()
            if card_details:
                other_parts.append(card_details)

    other = " | ".join(
        dict.fromkeys(part for part in other_parts if part)
    )

    return {
        "Konto": "comdirect",
        "Buchungstag": iso_date(
            record.get("Buchungstag (orig)", "")
        ),
        "Umsatzart": transaction_type,
        "Begünstigter / Auftraggeber": beneficiary,
        "Verwendungszweck": purpose,
        "Betrag": record.get("Ausgang/Eingang (orig)", ""),
        "Währung": "EUR",
        "Sonstiges": other,
    }

def write_csv(records, path, headers):
    """Write a German Excel-friendly CSV using semicolons."""
    with path.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.DictWriter(
            output_file,
            fieldnames=headers,
            delimiter=";",
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(records)


def main():
    script_dir = Path(__file__).resolve().parent
    pdf_files = sorted(
        (path for path in script_dir.iterdir()
         if path.is_file() and path.name.casefold().startswith("finanzreport")
         and path.suffix.casefold() == ".pdf"),
        key=lambda path: path.name.casefold(),
    )
    if not pdf_files:
        print(f"Keine Dateien passend zu 'Finanzreport*.pdf' gefunden in:\n{script_dir}")
        return 1

    records = []
    for pdf_path in pdf_files:
        print(f"Verarbeite: {pdf_path.name}")
        try:
            tables = parse_finanzreport(pdf_path)
            error = parse_error(tables, pdf_path.name)
            if error:
                print(f"WARNUNG: {error}")
                continue
            for account_title, table in tables.items():
                if not table.empty:
                    records.extend(prettify_rows(table, pdf_path.name, account_title))
        except Exception as exc:
            print(f"FEHLER bei {pdf_path.name}: {exc}")

    if not records:
        print("Es konnten keine Umsätze extrahiert werden; es wurde keine CSV geschrieben.")
        return 1

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_path = script_dir / f"Finanzreport_Umsaetze_{timestamp}.csv"
    suffix = 2
    while output_path.exists():
        output_path = script_dir / f"Finanzreport_Umsaetze_{timestamp}_{suffix}.csv"
        suffix += 1

    
    # Keep the existing detailed CSV.
    write_csv(records, output_path, OUTPUT_HEADERS)

    # Also generate the simplified CSV from the same transactions.
    simple_records = [
        make_simple_record(record)
        for record in records
    ]

    simple_path = script_dir / (
        f"Finanzreport_Umsaetze_einfach_{timestamp}.csv"
    )

    suffix = 2
    while simple_path.exists():
        simple_path = script_dir / (
            f"Finanzreport_Umsaetze_einfach_{timestamp}_{suffix}.csv"
        )
        suffix += 1

    write_csv(simple_records, simple_path, SIMPLE_HEADERS)

    print(f"Fertig: {len(records)} Umsätze aus {len(pdf_files)} PDF(s) verarbeitet.")
    print(f"Detaillierte CSV: {output_path.name}")
    print(f"Vereinfachte CSV: {simple_path.name}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
