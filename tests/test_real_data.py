"""
Abnahmetests mit den echten Listen aus den KARMA-Testfällen.

Die Dateien sind nicht versioniert. Ohne den Ordner testdaten/ werden die Tests
übersprungen. Die Kaufdatei hat 286 Zeilen; die Sollwerte im Wiki beruhen auf einer
von Hand bereinigten Fassung mit 283 Titeln (ohne die Zeilen 116, 147 und 264).
"""

from pathlib import Path

import pytest

from kbart_logic import compare_frames, find_isbn_columns, load_kbart, load_purchase

DATA = Path(__file__).resolve().parent.parent / "testdaten"
PURCHASE = DATA / "Cambridge_einzeltitel_Stand_2026_09_02.xlsx"
PURCHASE_283 = DATA / "Cambridge_einzeltitel_283.xlsx"
CURRENT = DATA / "CUP_Cambridge_Core_All_Books_2026-09-02.tsv"
WITHDRAWN = DATA / "CUP_Cambridge_Core_-_Withdrawn_Books_2026-09-02.tsv"
COMBINED = DATA / "CUP_Cambridge_Core__All_Books_including_withdrawn_titles_2026-09-02.tsv"
BIBLIOSCOUT = DATA / "BiblioScout_Global_AllBooks.tsv"

pytestmark = pytest.mark.skipif(not PURCHASE.exists(), reason="testdaten/ fehlt")


def run(path, purchase_path=PURCHASE):
    """Gleicht eine KBART-Datei über alle ISBN-Spalten mit der Kaufdatei ab."""
    purchase = load_purchase(purchase_path)
    return compare_frames(load_kbart(path), purchase, find_isbn_columns(purchase.columns))


def missing_isbns(result):
    """Liefert alle ISBNs aus der Liste fehlender Zeilen."""
    return {isbn for cell in result.missing["Nicht_gefundene_ISBNs"]
            for isbn in cell.split("; ")}


# Die vier im Wiki genannten, zurückgezogenen Titel
WITHDRAWN_ISBNS = {"9789048544653", "9789048544974", "9781139164726", "9780511800290"}


def test_current_list_misses_exactly_the_withdrawn_titles():
    """UC-05: Die vier zurückgezogenen Titel fehlen in der laufenden Liste."""
    assert WITHDRAWN_ISBNS <= missing_isbns(run(CURRENT))


def test_withdrawn_list_finds_the_withdrawn_titles():
    """UC-05: Die Withdrawn-Liste enthält genau die vier zurückgezogenen Titel."""
    result = run(WITHDRAWN)
    assert len(result.filtered) == 4
    assert not WITHDRAWN_ISBNS & missing_isbns(result)


def test_combined_list_contains_the_withdrawn_titles():
    """UC-05: Die Kombi-Liste enthält auch die zurückgezogenen Titel."""
    assert not WITHDRAWN_ISBNS & missing_isbns(run(COMBINED))


@pytest.mark.parametrize("path, hits, missing", [
    (CURRENT, 280, 5),
    (WITHDRAWN, 4, 281),
    (COMBINED, 284, 1),
])
def test_unfiltered_purchase_file_counts(path, hits, missing):
    """Die unbereinigte Kaufdatei mit 286 Zeilen liefert feste Ergebnisse."""
    result = run(path)
    assert (len(result.filtered), len(result.missing)) == (hits, missing)


@pytest.mark.skipif(not PURCHASE_283.exists(), reason="283er-Kaufdatei fehlt")
@pytest.mark.parametrize("path, expected_file, missing", [
    (CURRENT, "kbart_gefiltert_nurlaufendegeprüft.tsv", 4),
    (WITHDRAWN, "kbart_gefiltert_nurwithdrawngeprüft.tsv", 279),
    (COMBINED, "kbart_gefiltert_laufendundwithdrawn.tsv", 0),
])
def test_cleaned_purchase_file_matches_wiki_results(path, expected_file, missing):
    """Mit 283 Titeln entstehen genau die KBART-Dateien aus dem Wiki."""
    result = run(path, PURCHASE_283)
    expected = load_kbart(DATA / expected_file)
    assert sorted(result.filtered["online_identifier"]) == sorted(expected["online_identifier"])
    assert len(result.missing) == missing


@pytest.mark.skipif(not PURCHASE_283.exists(), reason="283er-Kaufdatei fehlt")
def test_cleaned_purchase_file_reports_wiki_line_numbers():
    """Die fehlenden Titel stehen in denselben Excel-Zeilen wie im Wiki."""
    result = run(CURRENT, PURCHASE_283)
    assert result.missing["Zeilennummer_Kaufdatei"].tolist() == [44, 67, 276, 282]


def test_unfiltered_purchase_file_notes():
    """UC-08: Die beiden Ocu-Titel der Kaufdatei werden gemeldet."""
    notes = run(CURRENT).notes
    assert notes["Zeilennummer_Kaufdatei"].tolist() == [147, 264]
    assert notes["Hinweis"].tolist() == [
        "Ocu: ISBN nicht in der KBART gefunden",
        "Prüfen: leere ISBN, Ocu oder bestellt? (BBG: Ocu)",
    ]


@pytest.mark.skipif(not PURCHASE_283.exists(), reason="283er-Kaufdatei fehlt")
def test_cleaned_purchase_file_has_no_notes():
    """UC-08: Nach der Bereinigung von Hand gibt es keine Hinweise mehr."""
    assert run(CURRENT, PURCHASE_283).notes.empty


def test_all_lines_of_biblioscout_are_read():
    """Anführungszeichen in Titeln dürfen keine Zeilen verschlucken."""
    assert len(load_kbart(BIBLIOSCOUT)) == 4617
