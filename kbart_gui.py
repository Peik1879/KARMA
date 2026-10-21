"""
KARMA - KBART Abgleich mit Titellisten.

Das Programm filtert eine KBART-Datei anhand der ISBNs einer Excel-Titelliste.
Übrig bleiben nur die Titel, die in beiden Dateien vorkommen; ISBNs ohne
Entsprechung in der KBART werden separat ausgewiesen.

Die Oberfläche basiert auf CustomTkinter. Die Auswertung läuft in einem
Hintergrund-Thread, damit das Fenster während des Abgleichs bedienbar bleibt.
"""

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk

import pandas as pd

try:
    import customtkinter as ctk
except ImportError as exc:  # pragma: no cover - nur ohne installierte Abhängigkeit
    sys.stderr.write(
        "Das Paket 'customtkinter' fehlt.\n"
        "Bitte installieren mit:  pip install -r requirements.txt\n"
    )
    raise SystemExit(1) from exc


# --------------------------------------------------------------------------
# Erscheinungsbild
# --------------------------------------------------------------------------

APP_NAME = "KARMA"
APP_SUBTITLE = "KBART-Abgleich mit Titellisten"

# Farbpaare jeweils als (heller Modus, dunkler Modus)
ACCENT = ("#0f6a7d", "#2b93a8")
ACCENT_HOVER = ("#0b5361", "#3aa9be")
SURFACE = ("#ffffff", "#232a2e")
SURFACE_SOFT = ("#f1f4f6", "#1c2226")
BORDER = ("#dde3e7", "#333c42")
TEXT_DIM = ("#66757e", "#93a1a8")
OK_COLOR = ("#2c6e3f", "#5fb87c")
WARN_COLOR = ("#9a5b12", "#e0a458")
ERROR_COLOR = ("#a52a2a", "#e57373")

PAD = 14

# Ab dieser Zeilenzahl wird die Vorschau gekürzt; der Export bleibt vollständig.
PREVIEW_LIMIT = 500

# Spalten, die in der KBART-Datei für den Abgleich vorhanden sein müssen.
REQUIRED_KBART_COLUMNS = ("online_identifier", "print_identifier")

# Diese KBART-Spalten werden in der Trefferliste angezeigt, sofern vorhanden.
PREVIEW_COLUMNS = (
    "publication_title",
    "print_identifier",
    "online_identifier",
    "publication_type",
    "date_monograph_published_online",
)


# --------------------------------------------------------------------------
# Auswertung (ohne Oberflächenbezug, dadurch einzeln testbar)
# --------------------------------------------------------------------------

def normalize_isbn(series):
    """
    Normalisiert ISBN-Werte und entfernt ungültige Einträge.

    Parameters:
        series (pd.Series): Serie mit ISBN-Werten.

    Returns:
        pd.Series: Serie mit bereinigten ISBNs ohne Leer- und Nullwerte.
    """
    return (
        series.astype(str)
        .str.replace("-", "", regex=False)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .replace("nan", pd.NA)
        .replace("0", pd.NA)
        .replace("", pd.NA)
        .dropna()
    )


def compare_frames(kbart_df, purchase_df, isbn_column, drop_serials=True):
    """
    Gleicht eine KBART-Tabelle gegen die ISBN-Spalte einer Titelliste ab.

    Parameters:
        kbart_df (pd.DataFrame): Eingelesene KBART-Datei.
        purchase_df (pd.DataFrame): Eingelesene Titelliste.
        isbn_column (str): Name der auszuwertenden Spalte in der Titelliste.
        drop_serials (bool): Zeitschriften vor dem Abgleich entfernen.

    Returns:
        tuple: (gefilterte KBART als DataFrame, Liste der fehlenden ISBNs)
    """
    if drop_serials and "publication_type" in kbart_df.columns:
        kbart_df = kbart_df[kbart_df["publication_type"] != "Serial"]

    purchase_isbns = normalize_isbn(purchase_df[isbn_column])

    online_isbns = normalize_isbn(kbart_df["online_identifier"])
    print_isbns = normalize_isbn(kbart_df["print_identifier"])
    kbart_isbns = set(online_isbns.unique()).union(set(print_isbns.unique()))

    online_matches = normalize_isbn(kbart_df["online_identifier"]).isin(
        purchase_isbns).reindex(kbart_df.index, fill_value=False)
    print_matches = normalize_isbn(kbart_df["print_identifier"]).isin(
        purchase_isbns).reindex(kbart_df.index, fill_value=False)
    filtered = kbart_df[online_matches | print_matches].copy()

    for column in REQUIRED_KBART_COLUMNS:
        filtered[column] = filtered[column].astype("string").str.replace(
            ".0", "", regex=False)

    missing = [
        str(isbn)
        for isbn in purchase_isbns[~purchase_isbns.isin(kbart_isbns)].tolist()
        if str(isbn) not in ("nan", "0", "")
    ]
    return filtered, missing


def format_number(value):
    """
    Formatiert eine Zahl mit Punkt als Tausendertrennzeichen.

    Parameters:
        value (int): Zu formatierende Zahl.

    Returns:
        str: Zahl als Text, zum Beispiel "3.482".
    """
    return f"{value:,}".replace(",", ".")


def display_value(value):
    """Formatiert leere Tabellenwerte als gut lesbaren Platzhalter."""
    if pd.isna(value):
        return "-"
    return str(value)


def human_size(path):
    """
    Gibt die Dateigröße in einer lesbaren Einheit zurück.

    Parameters:
        path (str): Pfad zur Datei.

    Returns:
        str: Größe als Text, zum Beispiel "1,8 MB".
    """
    try:
        size = os.path.getsize(path)
    except OSError:
        return "unbekannt"
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}".replace(".", ",")
        size /= 1024
    return f"{size:.0f} GB"


def guess_isbn_column(columns):
    """
    Schlägt anhand der Spaltennamen eine ISBN-Spalte vor.

    Parameters:
        columns (list): Liste der Spaltennamen.

    Returns:
        str: Der wahrscheinlichste Spaltenname oder der erste Eintrag.
    """
    for keyword in ("e-isbn", "eisbn", "online_identifier", "isbn"):
        for name in columns:
            if keyword in str(name).lower():
                return name
    return columns[0] if columns else ""


# --------------------------------------------------------------------------
# Oberfläche
# --------------------------------------------------------------------------

class FileCard(ctk.CTkFrame):
    """Auswahlkarte für eine Eingabedatei mit Namen, Kennzahlen und Button."""

    # pylint: disable=too-many-ancestors

    def __init__(self, master, title, hint, command):
        """
        Erzeugt die Karte.

        Parameters:
            master: Übergeordnetes Widget.
            title (str): Überschrift der Karte.
            hint (str): Erklärender Text unter der Überschrift.
            command (callable): Wird beim Klick auf den Button aufgerufen.
        """
        super().__init__(master, fg_color=SURFACE, border_width=1,
                         border_color=BORDER, corner_radius=6)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text=title, anchor="w",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
                         row=0, column=0, sticky="ew", padx=PAD, pady=(PAD, 0))

        self.hint_label = ctk.CTkLabel(self, text=hint, anchor="w", justify="left",
                                       text_color=TEXT_DIM, font=ctk.CTkFont(size=11))
        self.hint_label.grid(row=1, column=0, sticky="ew", padx=PAD, pady=(2, 8))

        self.name_label = ctk.CTkLabel(self, text="Keine Datei gewählt", anchor="w",
                                       justify="left", font=ctk.CTkFont(size=12))
        self.name_label.grid(row=2, column=0, sticky="ew", padx=PAD)

        self.meta_label = ctk.CTkLabel(self, text="", anchor="w", justify="left",
                                       text_color=TEXT_DIM, font=ctk.CTkFont(size=11))
        self.meta_label.grid(row=3, column=0, sticky="ew", padx=PAD, pady=(1, 0))

        self.button = ctk.CTkButton(self, text="Auswählen", width=110, height=30,
                                    command=command, fg_color=ACCENT,
                                    hover_color=ACCENT_HOVER)
        self.button.grid(row=4, column=0, sticky="w", padx=PAD, pady=(10, PAD))

    def set_file(self, name, meta, ok=True):
        """
        Zeigt die gewählte Datei an.

        Parameters:
            name (str): Dateiname.
            meta (str): Zusatzangaben wie Zeilenzahl und Größe.
            ok (bool): False färbt die Zusatzangabe als Fehler ein.
        """
        self.name_label.configure(text=name)
        self.meta_label.configure(text=meta, text_color=TEXT_DIM if ok else ERROR_COLOR)
        self.button.configure(text="Andere Datei")


class StatTile(ctk.CTkFrame):
    """Kachel für eine Kennzahl des Abgleichs."""

    # pylint: disable=too-many-ancestors

    def __init__(self, master, label, color=None):
        """
        Erzeugt die Kachel.

        Parameters:
            master: Übergeordnetes Widget.
            label (str): Beschriftung unter der Zahl.
            color: Farbpaar für die Zahl oder None für Standardfarbe.
        """
        super().__init__(master, fg_color=SURFACE_SOFT, corner_radius=6)
        self.value_label = ctk.CTkLabel(self, text="\u2013",
                                        font=ctk.CTkFont(size=24, weight="bold"))
        if color:
            self.value_label.configure(text_color=color)
        self.value_label.pack(anchor="w", padx=PAD, pady=(10, 0))
        ctk.CTkLabel(self, text=label, text_color=TEXT_DIM,
                     font=ctk.CTkFont(size=11)).pack(anchor="w", padx=PAD, pady=(0, 10))

    def set_value(self, value):
        """
        Setzt die angezeigte Zahl.

        Parameters:
            value (str): Anzuzeigender Text.
        """
        self.value_label.configure(text=value)


class KarmaApp(ctk.CTk):
    """Hauptfenster der Anwendung."""

    # pylint: disable=too-many-instance-attributes,too-many-ancestors

    def __init__(self):
        """Baut das Fenster und initialisiert den Zustand."""
        super().__init__()
        self.title(f"{APP_NAME} \u2013 {APP_SUBTITLE}")
        self.geometry("1020x820")
        self.minsize(900, 660)

        self.kbart_path = None
        self.purchase_path = None
        self.kbart_df = None
        self.purchase_df = None
        self.result_df = None
        self.missing = []
        self.events = queue.Queue()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        self._build_header()
        self._build_files()
        self._build_options()
        self._build_action()
        self._build_results()
        self._build_status()

        self._style_treeviews()
        self.after(120, self._poll_events)

    # -- Aufbau ------------------------------------------------------------

    def _build_header(self):
        """Erzeugt die Kopfzeile mit Titel und Umschalter für das Farbschema."""
        header = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0, height=64)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)
        header.grid_propagate(False)

        ctk.CTkLabel(header, text=APP_NAME,
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=ACCENT).grid(row=0, column=0, padx=(PAD + 4, 8), pady=14)
        ctk.CTkLabel(header, text=APP_SUBTITLE, text_color=TEXT_DIM,
                     font=ctk.CTkFont(size=12)).grid(row=0, column=1, sticky="w")

        self.mode_switch = ctk.CTkSwitch(header, text="Dunkel", width=48,
                                         command=self._toggle_mode,
                                         progress_color=ACCENT,
                                         font=ctk.CTkFont(size=11))
        self.mode_switch.grid(row=0, column=2, padx=PAD)

    def _build_files(self):
        """Erzeugt die beiden Dateikarten."""
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.grid(row=1, column=0, sticky="ew", padx=PAD, pady=(PAD, 0))
        row.grid_columnconfigure((0, 1), weight=1, uniform="files")

        self.kbart_card = FileCard(
            row, "1 \u00b7 KBART-Datei",
            "Tab-getrennte Datei des Anbieters (.txt oder .tsv)",
            self._choose_kbart)
        self.kbart_card.grid(row=0, column=0, sticky="nsew", padx=(0, 7))

        self.purchase_card = FileCard(
            row, "2 \u00b7 Titelliste",
            "Excel-Datei mit den erworbenen Titeln (.xlsx)",
            self._choose_purchase)
        self.purchase_card.grid(row=0, column=1, sticky="nsew", padx=(7, 0))

    def _build_options(self):
        """Erzeugt die Spaltenauswahl und die Zusatzoptionen."""
        box = ctk.CTkFrame(self, fg_color=SURFACE, border_width=1,
                           border_color=BORDER, corner_radius=6)
        box.grid(row=2, column=0, sticky="ew", padx=PAD, pady=PAD)
        box.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(box, text="3 \u00b7 ISBN-Spalte der Titelliste", anchor="w",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
                         row=0, column=0, columnspan=2, sticky="ew",
                         padx=PAD, pady=(PAD, 8))

        self.column_menu = ctk.CTkOptionMenu(
            box, values=["Erst Titelliste wählen"], width=260,
            command=self._on_column_change, state="disabled",
            fg_color=ACCENT, button_color=ACCENT_HOVER, button_hover_color=ACCENT)
        self.column_menu.grid(row=1, column=0, sticky="w", padx=PAD)

        self.preview_label = ctk.CTkLabel(
            box, text="Die Spalten werden nach dem Laden der Datei angeboten.",
            anchor="w", text_color=TEXT_DIM, font=ctk.CTkFont(size=11))
        self.preview_label.grid(row=1, column=1, sticky="w", padx=(PAD, PAD))

        self.serial_check = ctk.CTkCheckBox(
            box, text="Zeitschriften (publication_type = Serial) vor dem Abgleich entfernen",
            font=ctk.CTkFont(size=12), fg_color=ACCENT, hover_color=ACCENT_HOVER)
        self.serial_check.select()
        self.serial_check.grid(row=2, column=0, columnspan=2, sticky="w",
                               padx=PAD, pady=(12, PAD))

    def _build_action(self):
        """Erzeugt die Schaltfläche zum Start und den Fortschrittsbalken."""
        action_row = ctk.CTkFrame(self, fg_color="transparent")
        action_row.grid(row=3, column=0, sticky="ew", padx=PAD)
        action_row.grid_columnconfigure(1, weight=1)

        self.start_button = ctk.CTkButton(
            action_row, text="Abgleich starten", width=170, height=38,
            state="disabled", command=self._start_compare,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"))
        self.start_button.grid(row=0, column=0, sticky="w")

        self.progress = ctk.CTkProgressBar(action_row, height=5, progress_color=ACCENT)
        self.progress.grid(row=0, column=1, sticky="ew", padx=PAD)
        self.progress.set(0)

    def _build_results(self):
        """Erzeugt den Ergebnisbereich mit Kennzahlen, Tabellen und Export."""
        area = ctk.CTkFrame(self, fg_color="transparent")
        area.grid(row=4, column=0, sticky="nsew", padx=PAD, pady=PAD)
        area.grid_columnconfigure(0, weight=1)
        area.grid_rowconfigure(1, weight=1)

        tiles = ctk.CTkFrame(area, fg_color="transparent")
        tiles.grid(row=0, column=0, sticky="ew", pady=(0, PAD))
        tiles.grid_columnconfigure((0, 1, 2), weight=1, uniform="tiles")

        self.tile_hits = StatTile(tiles, "Titel in der neuen KBART", OK_COLOR)
        self.tile_hits.grid(row=0, column=0, sticky="ew", padx=(0, 7))
        self.tile_missing = StatTile(tiles, "ISBN ohne Treffer", WARN_COLOR)
        self.tile_missing.grid(row=0, column=1, sticky="ew", padx=7)
        self.tile_rate = StatTile(tiles, "Trefferquote der Titelliste")
        self.tile_rate.grid(row=0, column=2, sticky="ew", padx=(7, 0))

        self.tabs = ctk.CTkTabview(area, fg_color=SURFACE, corner_radius=6,
                                   segmented_button_selected_color=ACCENT,
                                   segmented_button_selected_hover_color=ACCENT_HOVER)
        self.tabs.grid(row=1, column=0, sticky="nsew")
        self.tabs.add("Gefundene Titel")
        self.tabs.add("Fehlende ISBN")

        self.hits_tree = self._make_tree(self.tabs.tab("Gefundene Titel"),
                                         ("Titel", "Print-ISBN", "Online-ISBN", "Typ"))
        self.missing_tree = self._make_tree(self.tabs.tab("Fehlende ISBN"),
                                            ("#", "ISBN aus der Titelliste"))

        exports = ctk.CTkFrame(area, fg_color="transparent")
        exports.grid(row=2, column=0, sticky="ew", pady=(PAD, 0))

        self.export_kbart_button = ctk.CTkButton(
            exports, text="Neue KBART speichern", width=190, height=34,
            state="disabled", command=self._export_kbart,
            fg_color=ACCENT, hover_color=ACCENT_HOVER)
        self.export_kbart_button.pack(side="left")

        self.export_missing_button = ctk.CTkButton(
            exports, text="Fehlende ISBN speichern", width=190, height=34,
            state="disabled", command=self._export_missing,
            fg_color="transparent", border_width=1, border_color=ACCENT,
            text_color=ACCENT, hover_color=SURFACE_SOFT)
        self.export_missing_button.pack(side="left", padx=8)

    def _build_status(self):
        """Erzeugt die Statuszeile am unteren Fensterrand."""
        self.status = ctk.CTkLabel(self, text="Bereit. Bitte beide Dateien wählen.",
                                   anchor="w", text_color=TEXT_DIM,
                                   font=ctk.CTkFont(size=11))
        self.status.grid(row=5, column=0, sticky="ew", padx=PAD + 2, pady=(0, 8))

    def _make_tree(self, parent, columns):
        """
        Legt eine Tabelle mit Bildlaufleiste an.

        Parameters:
            parent: Übergeordnetes Widget.
            columns (tuple): Spaltenüberschriften.

        Returns:
            ttk.Treeview: Die erzeugte Tabelle.
        """
        wrap = tk.Frame(parent, highlightthickness=0, bd=0)
        wrap.pack(fill="both", expand=True, padx=2, pady=2)

        tree = ttk.Treeview(wrap, columns=columns, show="headings",
                            selectmode="browse", style="Karma.Treeview")
        for name in columns:
            tree.heading(name, text=name)
            width = 90 if name in ("#", "Typ") else 240
            tree.column(name, width=width, anchor="w", stretch=True)

        scroll = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        return tree

    def _style_treeviews(self):
        """Passt das Aussehen der ttk-Tabellen an das Farbschema an."""
        dark = ctk.get_appearance_mode() == "Dark"
        bg = "#232a2e" if dark else "#ffffff"
        fg = "#e6ecef" if dark else "#1c242b"
        head_bg = "#1c2226" if dark else "#f1f4f6"
        sel = "#2b93a8" if dark else "#0f6a7d"

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Karma.Treeview", background=bg, fieldbackground=bg,
                        foreground=fg, rowheight=27, borderwidth=0,
                        font=("TkDefaultFont", 10))
        style.configure("Karma.Treeview.Heading", background=head_bg, foreground=fg,
                        relief="flat", font=("TkDefaultFont", 9, "bold"), padding=6)
        style.map("Karma.Treeview.Heading", background=[("active", head_bg)])
        style.map("Karma.Treeview", background=[("selected", sel)],
                  foreground=[("selected", "#ffffff")])

    # -- Interaktion -------------------------------------------------------

    def _toggle_mode(self):
        """Wechselt zwischen hellem und dunklem Farbschema."""
        ctk.set_appearance_mode("Dark" if self.mode_switch.get() else "Light")
        self._style_treeviews()

    def _set_status(self, text, color=None):
        """
        Aktualisiert die Statuszeile.

        Parameters:
            text (str): Anzuzeigender Text.
            color: Farbpaar oder None für die gedämpfte Standardfarbe.
        """
        self.status.configure(text=text, text_color=color or TEXT_DIM)

    def _choose_kbart(self):
        """Wählt die KBART-Datei aus und prüft ihre Pflichtspalten."""
        path = filedialog.askopenfilename(
            title="KBART-Datei wählen",
            filetypes=[("KBART / Tab-getrennt", "*.txt *.tsv *.csv"), ("Alle Dateien", "*.*")])
        if not path:
            return
        try:
            frame = pd.read_csv(path, sep="\t", dtype=str)
        except (OSError, UnicodeDecodeError, pd.errors.ParserError,
                pd.errors.EmptyDataError) as error:
            self.kbart_card.set_file(os.path.basename(path),
                                     f"Konnte nicht gelesen werden: {error}", ok=False)
            self._set_status("Die KBART-Datei konnte nicht geladen werden.", ERROR_COLOR)
            return

        fehlend = [c for c in REQUIRED_KBART_COLUMNS if c not in frame.columns]
        if fehlend:
            self.kbart_card.set_file(
                os.path.basename(path),
                "Pflichtspalte fehlt: " + ", ".join(fehlend), ok=False)
            self._set_status("Die Datei enthält nicht die erwarteten KBART-Spalten.",
                             ERROR_COLOR)
            self.kbart_df = None
            self._refresh_start_state()
            return

        self.kbart_path = path
        self.kbart_df = frame
        self.kbart_card.set_file(
            os.path.basename(path),
            f"{format_number(len(frame))} Zeilen \u00b7 {len(frame.columns)} Spalten "
            f"\u00b7 {human_size(path)}")
        self._set_status("KBART-Datei geladen.")
        self._refresh_start_state()

    def _choose_purchase(self):
        """Wählt die Titelliste aus und füllt die Spaltenauswahl."""
        path = filedialog.askopenfilename(
            title="Titelliste wählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xlsm *.xls"), ("Alle Dateien", "*.*")])
        if not path:
            return
        try:
            frame = pd.read_excel(path)
        except (OSError, ValueError, KeyError) as error:
            self.purchase_card.set_file(os.path.basename(path),
                                        f"Konnte nicht gelesen werden: {error}", ok=False)
            self._set_status("Die Titelliste konnte nicht geladen werden.", ERROR_COLOR)
            return

        if frame.empty or frame.columns.empty:
            self.purchase_card.set_file(os.path.basename(path),
                                        "Die Datei enthält keine Daten.", ok=False)
            return

        self.purchase_path = path
        self.purchase_df = frame
        self.purchase_card.set_file(
            os.path.basename(path),
            f"{format_number(len(frame))} Zeilen \u00b7 {len(frame.columns)} Spalten "
            f"\u00b7 {human_size(path)}")

        spalten = [str(c) for c in frame.columns]
        vorschlag = guess_isbn_column(spalten)
        self.column_menu.configure(values=spalten, state="normal")
        self.column_menu.set(vorschlag)
        self._on_column_change(vorschlag)
        self._set_status("Titelliste geladen. Bitte ISBN-Spalte prüfen.")
        self._refresh_start_state()

    def _on_column_change(self, column):
        """
        Zeigt Beispielwerte der gewählten Spalte an.

        Parameters:
            column (str): Name der gewählten Spalte.
        """
        if self.purchase_df is None or column not in [str(c) for c in self.purchase_df.columns]:
            return
        werte = normalize_isbn(self.purchase_df[column]).head(3).tolist()
        if werte:
            self.preview_label.configure(
                text="Beispiele: " + ", ".join(werte), text_color=TEXT_DIM)
        else:
            self.preview_label.configure(
                text="In dieser Spalte stehen keine auswertbaren ISBN.",
                text_color=WARN_COLOR)

    def _refresh_start_state(self):
        """Aktiviert die Startschaltfläche, sobald beide Dateien geladen sind."""
        bereit = self.kbart_df is not None and self.purchase_df is not None
        self.start_button.configure(state="normal" if bereit else "disabled")
        if bereit:
            self._set_status("Bereit für den Abgleich.")

    # -- Abgleich ----------------------------------------------------------

    def _start_compare(self):
        """Startet den Abgleich in einem Hintergrund-Thread."""
        self.start_button.configure(state="disabled", text="Abgleich läuft \u2026")
        self.export_kbart_button.configure(state="disabled")
        self.export_missing_button.configure(state="disabled")
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self._set_status("Die Dateien werden verglichen \u2026")

        column = self.column_menu.get()
        drop_serials = bool(self.serial_check.get())

        thread = threading.Thread(target=self._worker,
                                  args=(column, drop_serials), daemon=True)
        thread.start()

    def _worker(self, column, drop_serials):
        """
        Fuehrt den Abgleich aus und meldet das Ergebnis über die Queue.

        Parameters:
            column (str): Name der ISBN-Spalte.
            drop_serials (bool): Zeitschriften entfernen.
        """
        try:
            treffer, fehlend = compare_frames(self.kbart_df, self.purchase_df,
                                              column, drop_serials)
            self.events.put(("done", (treffer, fehlend)))
        except (KeyError, ValueError, TypeError) as error:  # pragma: no cover
            self.events.put(("error", str(error)))

    def _poll_events(self):
        """Holt Meldungen aus dem Hintergrund-Thread und aktualisiert die Oberfläche."""
        try:
            while True:
                art, nutzlast = self.events.get_nowait()
                if art == "done":
                    self._show_results(*nutzlast)
                elif art == "error":
                    self.progress.stop()
                    self.progress.set(0)
                    self.start_button.configure(state="normal", text="Abgleich starten")
                    self._set_status(f"Abbruch: {nutzlast}", ERROR_COLOR)
        except queue.Empty:
            pass
        self.after(120, self._poll_events)

    def _show_results(self, treffer, fehlend):
        """
        Stellt das Ergebnis in Kennzahlen und Tabellen dar.

        Parameters:
            treffer (pd.DataFrame): Gefilterte KBART-Zeilen.
            fehlend (list): ISBN ohne Entsprechung in der KBART.
        """
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1)
        self.start_button.configure(state="normal", text="Abgleich starten")

        self.result_df = treffer
        self.missing = fehlend

        gesamtzahl = len(treffer) + len(fehlend)
        quote = (len(treffer) / gesamtzahl * 100) if gesamtzahl else 0
        self.tile_hits.set_value(format_number(len(treffer)))
        self.tile_missing.set_value(format_number(len(fehlend)))
        self.tile_rate.set_value(f"{quote:.1f} %".replace(".", ","))

        self.hits_tree.delete(*self.hits_tree.get_children())
        spalten = [c for c in PREVIEW_COLUMNS if c in treffer.columns]
        for _, zeile in treffer.head(PREVIEW_LIMIT).iterrows():
            self.hits_tree.insert("", "end", values=(
                display_value(zeile.get("publication_title", "")),
                display_value(zeile.get("print_identifier", "")),
                display_value(zeile.get("online_identifier", "")),
                display_value(zeile.get("publication_type", "")),
            ))

        self.missing_tree.delete(*self.missing_tree.get_children())
        for nummer, isbn in enumerate(fehlend[:PREVIEW_LIMIT], start=1):
            self.missing_tree.insert("", "end", values=(nummer, isbn))

        self.export_kbart_button.configure(state="normal" if len(treffer) else "disabled")
        self.export_missing_button.configure(state="normal" if fehlend else "disabled")

        hinweis = ""
        if len(treffer) > PREVIEW_LIMIT or len(fehlend) > PREVIEW_LIMIT:
            hinweis = (f" Die Vorschau zeigt die ersten {PREVIEW_LIMIT} Zeilen, "
                       "der Export enthält alle.")
        self._set_status(
            f"Abgleich abgeschlossen \u00b7 {len(spalten)} KBART-Spalten übernommen."
            + hinweis, OK_COLOR)

    # -- Export ------------------------------------------------------------

    def _export_kbart(self):
        """Speichert die gefilterte KBART-Datei."""
        if self.result_df is None or self.result_df.empty:
            return
        pfad = filedialog.asksaveasfilename(
            defaultextension=".tsv", filetypes=[("KBART / TSV", "*.tsv")],
            title="Neue KBART-Datei speichern")
        if not pfad:
            return
        try:
            self.result_df.to_csv(pfad, sep="\t", index=False, encoding="utf-8")
        except OSError as error:
            self._set_status(f"Speichern fehlgeschlagen: {error}", ERROR_COLOR)
            return
        self._set_status(f"Gespeichert: {pfad}", OK_COLOR)

    def _export_missing(self):
        """Speichert die Liste der ISBN ohne Treffer."""
        if not self.missing:
            return
        pfad = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Textdatei", "*.txt")],
            title="Fehlende ISBN speichern")
        if not pfad:
            return
        try:
            with open(pfad, "w", encoding="utf-8") as datei:
                datei.write("\n".join(self.missing))
        except OSError as error:
            self._set_status(f"Speichern fehlgeschlagen: {error}", ERROR_COLOR)
            return
        self._set_status(f"Gespeichert: {pfad}", OK_COLOR)


def main():
    """Startet die Anwendung."""
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")
    app = KarmaApp()
    app.mainloop()


if __name__ == "__main__":
    main()
