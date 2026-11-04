"""Tests der Auswertung anhand der Beispiele aus den KARMA-Testfällen."""

import pandas as pd
import pytest

from kbart_logic import (
    compare_frames,
    find_isbn_columns,
    hit_rate,
    load_kbart,
    load_purchase,
    normalize_identifier,
    split_identifiers,
    write_tsv,
)

KBART_HEADER = ["publication_title", "print_identifier", "online_identifier",
                "publication_type"]


def kbart(*rows):
    """Baut eine KBART-Tabelle aus Zeilen im Format von KBART_HEADER."""
    return pd.DataFrame(list(rows), columns=KBART_HEADER)


@pytest.mark.parametrize("raw, expected", [
    ("978-3-8047-4377-9", "9783804743779"),
    ("9783804743779.0", "9783804743779"),
    ("978 3 8047 4377 9", "9783804743779"),
    ("978‑3‑8047‑4377‑9", "9783804743779"),
    ("978-0-521-86590-6$fhardback", "9780521865906"),
    ("9781009597371fhardback", "9781009597371"),
    ("080442957x", "080442957X"),
    ("0-8044-2957-X", "080442957X"),
    ("10.33196/9783704699909", "10.33196/9783704699909"),
    ("10.0000/abc", "10.0000/abc"),
    ("  ", ""),
    ("0", ""),
    ("nan", ""),
    (float("nan"), ""),
])
def test_normalize_identifier(raw, expected):
    """UC-03: Normalisierung ohne Beschädigung anderer Identifikatoren."""
    assert normalize_identifier(raw) == expected


def test_split_identifiers_removes_blanks_and_duplicates():
    """UC-03: Mehrere Werte in einer Zelle werden einzeln ausgewertet."""
    value = "978-1-108-47098-8; 978-1-108-45691-3; 978-1-108-47098-8;"
    assert split_identifiers(value) == ["9781108470988", "9781108456913"]


def test_find_isbn_columns():
    """UC-08: Alle Spalten mit ISBN im Namen werden erkannt."""
    columns = ["PPN", "ISBN", "ISBN_1", "isbn_2", "Titel"]
    assert find_isbn_columns(columns) == ["ISBN", "ISBN_1", "isbn_2"]


def test_empty_identifiers_do_not_stop_the_comparison():
    """UC-02: Zeilen ohne Identifier werden nicht zum Abbruch und nicht zum Treffer."""
    data = kbart(["leer", "", "", "Book"], ["A", "9781111111111", "", "Book"])
    purchase = pd.DataFrame({"ISBN": ["9781111111111"]})
    result = compare_frames(data, purchase, ["ISBN"])
    assert result.filtered["publication_title"].tolist() == ["A"]
    assert result.missing.empty


@pytest.mark.parametrize("kbart_value, purchase_value", [
    ("9781111111111;9782222222222", "9782222222222"),
    ("9782222222222", "9781111111111; 9782222222222"),
])
def test_semicolons_match_on_both_sides(kbart_value, purchase_value):
    """UC-03: Semikolon-Listen treffen in KBART und in Titelliste."""
    data = kbart(["S", kbart_value, "", "Book"])
    purchase = pd.DataFrame({"ISBN": [purchase_value]})
    result = compare_frames(data, purchase, ["ISBN"])
    assert len(result.filtered) == 1
    assert result.missing.empty


def test_hit_in_any_isbn_column_counts_for_the_whole_row():
    """Eine Zeile ist gefunden, wenn eine ihrer Spalten trifft."""
    data = kbart(["B", "9782222222222", "", "Book"])
    purchase = pd.DataFrame({"ISBN": ["9781111111111", "9783333333333"],
                             "ISBN_2": ["9782222222222", None]})
    result = compare_frames(data, purchase, ["ISBN", "ISBN_2"])
    assert len(result.filtered) == 1
    assert result.missing.values.tolist() == [[3, "9783333333333"]]


def test_missing_rows_use_excel_line_numbers_and_skip_empty_rows():
    """Fehlende Titel tragen die Excel-Zeile; Zeilen ohne ISBN zählen nicht."""
    data = kbart(["A", "9781111111111", "", "Book"])
    purchase = pd.DataFrame({"ISBN": ["9781111111111", None, "978-9-999-99999-9"]})
    result = compare_frames(data, purchase, ["ISBN"])
    assert result.missing.values.tolist() == [[4, "9789999999999"]]
    assert result.rows_with_isbn == 2
    assert hit_rate(result) == 50


def test_duplicate_purchase_rows_do_not_distort_the_hit_rate():
    """Die Trefferquote bezieht sich auf Zeilen der Titelliste."""
    data = kbart(["A", "9781111111111", "", "Book"])
    purchase = pd.DataFrame({"ISBN": ["9781111111111", "9781111111111"]})
    assert hit_rate(compare_frames(data, purchase, ["ISBN"])) == 100


def test_hit_rate_without_isbn_is_zero():
    """Ohne auswertbare ISBN gibt es keine Division durch null."""
    data = kbart(["A", "9781111111111", "", "Book"])
    purchase = pd.DataFrame({"ISBN": [None]})
    assert hit_rate(compare_frames(data, purchase, ["ISBN"])) == 0


def notes_for(purchase, columns=("ISBN",), kbart_isbn="9781111111111"):
    """Gleicht eine Titelliste ab und liefert ihre Hinweise als Listen."""
    data = kbart(["A", kbart_isbn, "", "Book"])
    return compare_frames(data, pd.DataFrame(purchase), list(columns)).notes.values.tolist()


def test_row_without_any_isbn_gets_the_wiki_note():
    """UC-08: Ohne ISBN in allen Spalten erscheint die Meldung aus dem Testfall."""
    notes = notes_for({"Titel+Zusatz": ["Altdaten"], "ISBN": [None], "ISBN_1": [None]},
                      ("ISBN", "ISBN_1"))
    assert notes == [[2, "Altdaten", "Prüfen: leere ISBN, Ocu oder bestellt?"]]


def test_empty_row_note_names_the_ocu_code():
    """UC-08: Bei leerer ISBN wird die vorhandene Gattung mit ausgegeben."""
    notes = notes_for({"Titel+Zusatz": ["T"], "ISBN": [None], "BBG": ["Ocu"]})
    assert notes == [[2, "T", "Prüfen: leere ISBN, Ocu oder bestellt? (BBG: Ocu)"]]


def test_any_filled_isbn_column_prevents_the_empty_note():
    """UC-08: Eine gefüllte ISBN-Spalte genügt, auch wenn sie nicht gewählt ist."""
    notes = notes_for({"Titel+Zusatz": ["T"], "ISBN": [None], "ISBN_2": ["9781111111111"]})
    assert notes == []


@pytest.mark.parametrize("code", ["Ocu", "ocu", " OCU "])
def test_ocu_with_isbn_reports_whether_it_was_found(code):
    """UC-08: Auch Ocu mit ISBN werden gekennzeichnet, unabhängig von der Schreibweise."""
    notes = notes_for({"Titel+Zusatz": ["Gefunden", "Fehlt"],
                       "ISBN": ["9781111111111", "9782222222222"],
                       "BBG": [code, code]})
    assert notes == [[2, "Gefunden", "Ocu: ISBN in der KBART gefunden, bitte prüfen"],
                     [3, "Fehlt", "Ocu: ISBN nicht in der KBART gefunden"]]


def test_regular_titles_and_blank_rows_get_no_note():
    """Normale Titel und vollständig leere Excel-Zeilen erzeugen keinen Hinweis."""
    notes = notes_for({"Titel+Zusatz": ["T", None], "ISBN": ["9781111111111", None],
                       "BBG": ["Oau", None]})
    assert notes == []


def test_notes_work_without_title_and_bbg_columns():
    """Titelliste ohne Titel- und BBG-Spalte führt nicht zum Abbruch."""
    assert notes_for({"ISBN": [None], "PPN": ["1"]}) == [
        [2, "", "Prüfen: leere ISBN, Ocu oder bestellt?"]]


def test_serials_are_only_dropped_on_request():
    """Der Zeitschriftenfilter bleibt abschaltbar."""
    data = kbart(["Z", "", "9782222222222", "Serial"])
    purchase = pd.DataFrame({"ISBN": ["9782222222222"]})
    assert compare_frames(data, purchase, ["ISBN"]).filtered.empty
    assert len(compare_frames(data, purchase, ["ISBN"], drop_serials=False).filtered) == 1


def test_export_removes_float_suffix_but_keeps_other_values():
    """UC-03: Nur ein angehängtes .0 verschwindet, DOIs und Bindestriche bleiben."""
    data = kbart(["F", "9781234567890.0", "10.0000/abc", "Book"],
                 ["G", "978-1-234-56789-7", "9781234567897", "Book"])
    purchase = pd.DataFrame({"ISBN": ["9781234567890", "9781234567897"]})
    filtered = compare_frames(data, purchase, ["ISBN"]).filtered
    assert filtered["print_identifier"].tolist() == ["9781234567890", "978-1-234-56789-7"]
    assert filtered["online_identifier"].tolist() == ["10.0000/abc", "9781234567897"]


def test_compare_does_not_change_its_input():
    """UC-07: Eingabetabellen bleiben unverändert."""
    data = kbart(["F", "9781234567890.0", "", "Book"])
    before = data.copy()
    compare_frames(data, pd.DataFrame({"ISBN": ["9781234567890"]}), ["ISBN"])
    pd.testing.assert_frame_equal(data, before)


def test_quotes_and_na_values_survive_reading_and_writing(tmp_path):
    """Anführungszeichen verschmelzen keine Zeilen und NA bleibt Text."""
    source = tmp_path / "kbart.tsv"
    content = ("publication_title\tprint_identifier\tonline_identifier\n"
               '"Offenes Zitat\t9781\t\n'
               'Mit "Zitat" im Titel\t9782\t\n'
               "NA\t9783\t\n")
    source.write_text(content, encoding="utf-8")

    frame = load_kbart(source)
    assert len(frame) == 3
    assert frame["publication_title"].tolist()[2] == "NA"

    target = tmp_path / "neu.tsv"
    write_tsv(frame, target)
    assert target.read_text(encoding="utf-8") == content


def test_load_kbart_removes_byte_order_mark(tmp_path):
    """Eine Datei mit BOM behält den richtigen Namen der ersten Spalte."""
    source = tmp_path / "bom.tsv"
    source.write_text("publication_title\tprint_identifier\nA\t1\n", encoding="utf-8-sig")
    assert load_kbart(source).columns[0] == "publication_title"


def test_load_purchase_reads_numbers_without_float_suffix(tmp_path):
    """Als Zahl gespeicherte ISBNs ergeben nach dem Einlesen die richtige ISBN."""
    path = tmp_path / "kauf.xlsx"
    pd.DataFrame({"ISBN": [9780521129374, None], 2024: ["a", "b"]}).to_excel(path, index=False)
    frame = load_purchase(path)
    assert list(frame.columns) == ["ISBN", "2024"]
    assert split_identifiers(frame["ISBN"].iloc[0]) == ["9780521129374"]
