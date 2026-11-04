"""
KARMA - Auswertung ohne Oberflächenbezug.

Das Modul liest KBART-Dateien und Titellisten, bereinigt Identifikatoren und
gleicht beide Dateien ab. Es kommt ohne Fenster aus und ist dadurch einzeln testbar.
"""

import csv
import re
from typing import NamedTuple

import pandas as pd


# Spalten, die in der KBART-Datei für den Abgleich vorhanden sein müssen.
REQUIRED_KBART_COLUMNS = ("online_identifier", "print_identifier")

# Spalten der Liste fehlender Titel (Format der bisherigen Berichte).
MISSING_COLUMNS = ("Zeilennummer_Kaufdatei", "Nicht_gefundene_ISBNs")

EMPTY_VALUES = {"", "0", "nan", "none", "<na>"}

# Leerraum, Bindestrich, weiches Trennzeichen und typografische Striche
SEPARATORS = re.compile(r"[\s\-­‐-―]")
FLOAT_SUFFIX = re.compile(r"\.0$")

# ISBN am Wertanfang; Anhänge wie "$fhardback" oder "fhardback" dürfen folgen.
ISBN_PREFIX = re.compile(r"(\d{13}|\d{9}[\dXx])(?![\dXx])")

# Ganze Zahl oder ISBN mit Bindestrichen, die Excel zur Kommazahl gemacht hat
FLOAT_IDENTIFIER = re.compile(r"[\dXx\-]+\.0")


class CompareResult(NamedTuple):
    """
    Ergebnis eines Abgleichs.

    Attributes:
        filtered (pd.DataFrame): KBART-Zeilen, die zur Titelliste passen.
        missing (pd.DataFrame): Zeilen der Titelliste ohne Treffer in der KBART.
        rows_with_isbn (int): Anzahl der Zeilen der Titelliste mit mindestens einer ISBN.
    """

    filtered: pd.DataFrame
    missing: pd.DataFrame
    rows_with_isbn: int


def normalize_identifier(value):
    """
    Bereinigt einen einzelnen Identifikator.

    ISBNs werden auf Ziffern und ein großes X reduziert. Andere Werte wie DOIs
    bleiben bis auf äußeren Leerraum unverändert.

    Parameters:
        value: Wert aus einer Tabellenzelle.

    Returns:
        str: Bereinigter Wert oder ein leerer Text, wenn die Zelle leer ist.
    """
    if pd.isna(value):
        return ""
    text = str(value).strip()
    cleaned = FLOAT_SUFFIX.sub("", SEPARATORS.sub("", text))
    if cleaned.lower() in EMPTY_VALUES:
        return ""
    match = ISBN_PREFIX.match(cleaned)
    return match.group(1).upper() if match else text


def split_identifiers(value):
    """
    Zerlegt eine Zelle an Semikolons in bereinigte Identifikatoren.

    Parameters:
        value: Wert aus einer Tabellenzelle.

    Returns:
        list: Bereinigte Werte ohne Leereinträge und Dopplungen.
    """
    if pd.isna(value):
        return []
    parts = (normalize_identifier(part) for part in str(value).split(";"))
    return list(dict.fromkeys(part for part in parts if part))


def find_isbn_columns(columns):
    """
    Sucht alle Spalten, die "ISBN" im Namen tragen.

    Parameters:
        columns (iterable): Spaltennamen.

    Returns:
        list: Passende Spaltennamen in der ursprünglichen Reihenfolge.
    """
    return [name for name in columns if "isbn" in str(name).lower()]


def row_identifiers(frame, columns):
    """
    Sammelt je Tabellenzeile die Identifikatoren der angegebenen Spalten.

    Parameters:
        frame (pd.DataFrame): Auszuwertende Tabelle.
        columns (iterable): Zu lesende Spalten.

    Returns:
        list: Eine Liste bereinigter Werte je Zeile.
    """
    rows = [[] for _ in range(len(frame))]
    for column in columns:
        for row, value in zip(rows, frame[column]):
            row.extend(split_identifiers(value))
    return [list(dict.fromkeys(row)) for row in rows]


def clean_float_identifier(value):
    """
    Entfernt ein angehängtes ".0" von Identifikatoren, die als Zahl gelesen wurden.

    Parameters:
        value: Wert aus einer Tabellenzelle.

    Returns:
        Wert ohne ".0"; andere Werte bleiben unverändert.
    """
    if isinstance(value, str) and FLOAT_IDENTIFIER.fullmatch(value.strip()):
        return FLOAT_SUFFIX.sub("", value.strip())
    return value


def compare_frames(kbart_df, purchase_df, isbn_columns, drop_serials=True):
    """
    Gleicht eine KBART-Tabelle gegen die ISBN-Spalten einer Titelliste ab.

    Eine Zeile der Titelliste gilt als gefunden, sobald eine ihrer ISBNs in der
    KBART vorkommt.

    Parameters:
        kbart_df (pd.DataFrame): Eingelesene KBART-Datei.
        purchase_df (pd.DataFrame): Eingelesene Titelliste.
        isbn_columns (list): Auszuwertende Spalten der Titelliste.
        drop_serials (bool): Zeitschriften vor dem Abgleich entfernen.

    Returns:
        CompareResult: Treffer, fehlende Zeilen und Anzahl der geprüften Zeilen.
    """
    if drop_serials and "publication_type" in kbart_df.columns:
        kbart_df = kbart_df[kbart_df["publication_type"] != "Serial"]

    purchase_rows = row_identifiers(purchase_df, isbn_columns)
    purchase_isbns = {isbn for row in purchase_rows for isbn in row}

    kbart_rows = row_identifiers(kbart_df, REQUIRED_KBART_COLUMNS)
    kbart_isbns = {isbn for row in kbart_rows for isbn in row}

    mask = pd.Series([not purchase_isbns.isdisjoint(row) for row in kbart_rows],
                     index=kbart_df.index, dtype=bool)
    filtered = kbart_df[mask].copy()
    for column in REQUIRED_KBART_COLUMNS:
        filtered[column] = filtered[column].map(clean_float_identifier)

    # +2: Excel zählt ab 1 und die erste Zeile enthält die Überschriften.
    missing = pd.DataFrame(
        [
            (position + 2, "; ".join(row))
            for position, row in enumerate(purchase_rows)
            if row and kbart_isbns.isdisjoint(row)
        ],
        columns=MISSING_COLUMNS,
    )
    rows_with_isbn = sum(1 for row in purchase_rows if row)
    return CompareResult(filtered, missing, rows_with_isbn)


def hit_rate(result):
    """
    Berechnet den Anteil der gefundenen Zeilen der Titelliste.

    Parameters:
        result (CompareResult): Ergebnis eines Abgleichs.

    Returns:
        float: Anteil in Prozent; 0, wenn keine Zeile eine ISBN enthält.
    """
    if not result.rows_with_isbn:
        return 0.0
    found = result.rows_with_isbn - len(result.missing)
    return found / result.rows_with_isbn * 100


def load_kbart(path):
    """
    Liest eine KBART-Datei buchstabengetreu ein.

    Anführungszeichen gelten als normale Zeichen, weil KBART kein Quoting kennt.
    Sonst können Zeilen verschmelzen. Werte wie "NA" bleiben Text.

    Parameters:
        path (str): Pfad zur tabulatorgetrennten Datei.

    Returns:
        pd.DataFrame: Tabelle, in der jede Zelle Text ist.
    """
    frame = pd.read_csv(path, sep="\t", dtype=str, encoding="utf-8-sig",
                        quoting=csv.QUOTE_NONE, na_filter=False)
    return frame.fillna("")


def load_purchase(path):
    """
    Liest das erste Tabellenblatt einer Titelliste als Text ein.

    Parameters:
        path (str): Pfad zur Excel-Datei.

    Returns:
        pd.DataFrame: Tabelle mit Spaltennamen und Zellen als Text.
    """
    frame = pd.read_excel(path, dtype=str)
    frame.columns = [str(name) for name in frame.columns]
    return frame


def write_tsv(frame, path):
    """
    Schreibt eine Tabelle ohne Quoting als tabulatorgetrennte UTF-8-Datei.

    Parameters:
        frame (pd.DataFrame): Zu schreibende Tabelle.
        path (str): Zielpfad.
    """
    with open(path, "w", encoding="utf-8", newline="") as target:
        target.write("\t".join(str(name) for name in frame.columns) + "\n")
        for row in frame.itertuples(index=False):
            cells = ("" if pd.isna(cell) else str(cell) for cell in row)
            target.write("\t".join(cells) + "\n")
