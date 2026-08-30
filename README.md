# Compare-ExcelFiles — User Guide / Instrukcja obsługi / Bedienungsanleitung

[English](#english-user-guide) | [Polski](#polska-instrukcja-obsługi) | [Deutsch](#deutsche-bedienungsanleitung)

---

<a name="english-user-guide"></a>
# English User Guide

## 📌 Overview

**`Compare-ExcelFiles.ps1`** is a high-performance PowerShell WPF desktop application designed for comparing two Excel (`.xlsx`, `.xls`) or delimited text (`.csv`) files with mismatched column structures, different naming conventions, combined/concatenated fields, or differing layouts.

### 🌟 Key Capabilities
- **🌙 Dynamic Dark & Light Theme Modes**:
  - Full Dark Mode support (`#0F172A` deep slate background, `#1E293B` cards, `#334155` inputs/headers, `#F8FAFC` crisp high-contrast text).
  - Native Windows 10/11 DWM dark title bar integration via `DwmSetWindowAttribute`.
  - Contrast-optimized row status colors (Emerald for Added, Crimson for Deleted, Dark Amber for Modified) and vibrant gold cell diff highlighting (`#B45309` / `#FBBF24`).
  - Real-time theme toggle (`🌙 Dark` / `☀️ Light`) directly from the Setup Wizard header, Results Viewer toolbar, or Options tab.
  - Theme preference is automatically saved to `%APPDATA%\Compare_ExcelFiles\settings.xml`.
- **⚡ 1-Click Automatic Column Mapping (Auto-Map)**:
  - Automatically matches identical/similar columns between Base and Update files (case-insensitive & trimmed).
  - Smart Join Key detection: automatically pre-selects primary identifiers (`ID`, `Lp`, `Kod`, `Nr`, `Name`, `Imię i nazwisko`) in Tab 3.
  - Triggers automatically upon worksheet selection or on demand via `⚡ Auto-Map Columns`.
- **🔍 Intelligent Field Normalization & Punctuation Invariance**:
  - **Punctuation & Special Character Invariance**: Strips commas, periods, dashes, slashes, and symbols during comparison so that `"Street 5, City"` matches `"Street 5 - City"`.
  - **Whitespace Invariance**: Collapses all multiple/irregular spaces so formatting differences don't create false positives.
  - **Unicode Compatibility**: Full support for international diacritics (`ą, ć, ę, ł, ń, ó, ś, ź, ż, ä, ö, ü, ß`) and numbers.
- **✨ Cell-Level Diff Highlighting**:
  - In modified rows, only the cells that actually changed are highlighted with a distinct background and border.
- **🔄 Dual Layout Switcher (Top / Bottom vs. Side-by-Side)**:
  - Defaults to **Top / Bottom (Vertical)** layout, ideal for wide data tables with many columns.
  - Easily toggle between Top/Bottom and Side-by-Side views on the fly.
- **🔀 Asymmetric Column Mapping (N:1 or 1:N)**:
  - Merge multiple base columns (e.g., `First Name` + `Last Name`) to compare against a single update column (`Full Name`).
  - Merge Modes: `Concatenate` (with custom separator), `Exact`, and `FirstNonEmpty`.
- **🔑 Composite Join Keys**:
  - Match rows using single or multi-column composite keys (e.g., `Branch` + `Account Number`).
- **🌐 Trilingual UI**:
  - Complete localization in English (EN), Polish (PL), and German (DE).
- **📊 Formatted Excel Export**:
  - Export diff results to a clean Excel spreadsheet (`.xlsx`) with status flags and changed column lists.

---

## 💻 Prerequisites & Installation

- **Operating System**: Windows 10 / Windows 11 / Windows Server 2016+.
- **PowerShell**: Windows PowerShell 5.1 or PowerShell Core 7+.
- **Module**: `ImportExcel` (installed/imported automatically on launch if missing).

---

## 🚀 How to Run

Launch the script in PowerShell:

```powershell
& 'D:\Skrypty\Mnich_Adam_Skrypty\Compare-ExcelFiles.ps1'
```

---

## 📖 Step-by-Step Instructions

### Step 1: Select Files & Worksheets (Tab 1: Files)
1. Under **Base File**, click **"Select Base File"** and choose your reference `.xlsx` / `.csv` file.
2. Select the target worksheet from the **Sheet** dropdown.
3. Under **Update File**, click **"Select Update File"** and choose the modified `.xlsx` / `.csv` file.
4. Select the corresponding worksheet from the **Sheet** dropdown.

### Step 2: Define Column Mappings (Tab 2: Column Mapping)
- **Automatic**: Click **`⚡ Auto-Map Columns`** to generate 1:1 mappings for all matching column names.
- **Manual / Concatenated**: Click **"Add Mapping Rule"** to define custom multi-column relationships.
  - Select one or more columns from the Base list (use Ctrl+Click for multiple).
  - Select one or more columns from the Update list.
  - Choose merge mode (`Concatenate`, `Exact`, `FirstNonEmpty`) and separator (e.g., ` | ` or space).
  - Review live preview and click **OK**.
- **Edit / Remove**: Select a rule from the table and click **"Edit Selected"** or **"Remove Selected"**.

### Step 3: Select Join Key (Tab 3: Join Key)
- Select the column(s) that uniquely identify each row (e.g., `ID`, `PESEL`, `Code`, `EmployeeNumber`).
- If you ran Auto-Map, matching ID columns are automatically pre-selected for you.

### Step 4: Configure Comparison Options (Tab 4: Options)
- **Layout**: Select `Top / Bottom (Vertical)` [Recommended for wide tables] or `Side by Side (Horizontal)`.
- **Theme**: Select `☀️ Light` or `🌙 Dark`.
- **Ignore case when comparing**: Treats uppercase and lowercase letters as identical.
- **Trim whitespace before comparing**: Strips leading and trailing spaces.
- **Ignore punctuation & special characters**: Removes commas, hyphens, slashes, etc.
- **Ignore all internal whitespace differences**: Collapses internal spaces in merged fields.
- **Hide Unchanged rows in results**: Filters the view to show only Added, Deleted, and Modified rows.

### Step 5: Execute & Analyze (Results Viewer)
1. Click **"▶ Run Comparison"**.
2. A progress indicator will display while parsing and comparing data.
3. The **Diff Results Viewer** opens maximized:
   - **Status Bar Summary**: Total rows, Added count, Deleted count, Modified count, Unchanged count.
   - **Filter Toolbar**:
     - Filter by status (`(All)`, `Added`, `Deleted`, `Modified`, `Unchanged`).
     - Real-time search box across all columns.
     - Switch view layout (`Top / Bottom` or `Side by Side`).
     - Toggle Theme (`🌙 Dark` / `☀️ Light`).
     - **Sync Scroll**: Keeps base and update grids scrolled to the exact same position simultaneously.
   - **Export**: Click **"Export"** to save the combined diff report to a new Excel file.

---
---

<a name="polska-instrukcja-obsługi"></a>
# Polska Instrukcja Obsługi

## 📌 Przegląd narzędzia

**`Compare-ExcelFiles.ps1`** to zaawansowana aplikacja PowerShell z interfejsem graficznym WPF przeznaczona do precyzyjnego porównywania dwóch plików Excel (`.xlsx`, `.xls`) lub `.csv`, które mogą różnić się strukturą kolumn, nazwami nagłówków, układem pól lub zawierać dane połączone/scalone.

### 🌟 Główne możliwości
- **🌙 Dynamiczny Tryb Ciemny (Dark Mode) i Jasny (Light Mode)**:
  - Pełne wsparcie dla Dark Mode (tło `#0F172A`, karty `#1E293B`, kontrolki `#334155`, wysoki kontrast tekstu `#F8FAFC`).
  - Natywna integracja z ciemnym paskiem tytułowym Windows 10/11 poprzez API DWM (`DwmSetWindowAttribute`).
  - Zoptymalizowane kolory statusów wierszy (szmaragdowy dla dodanych, karmazynowy dla usuniętych, ciemnobrązowy dla zmodyfikowanych) oraz złoto-bursztynowe wyróżnienie zmienionych komórek (`#B45309` / `#FBBF24`).
  - Przełącznik motywu dostępny na pasku nagłówka kreatora, pasku narzędzi przeglądarki wyników oraz w zakładce Opcje.
  - Wybór motywu i języka jest automatycznie zapisywany w pliku `%APPDATA%\Compare_ExcelFiles\settings.xml`.
- **⚡ 1-Klikowe Automatyczne Mapowanie Kolumn (Auto-Map)**:
  - Automatyczne tworzenie reguł mapowania 1:1 dla pasujących kolumn bez konieczności ręcznego klikania.
  - Inteligentne wykrywanie klucza łączenia: automatyczne zaznaczanie kolumn identyfikacyjnych (`ID`, `Lp`, `Kod`, `Nr`, `Imię i nazwisko`) w Zakładce 3.
  - Uruchamia się samoczynnie po wyborze arkuszy lub na żądanie przyciskiem `⚡ Automatyczne mapowanie`.
- **🔍 Inteligentna normalizacja pól i odporność na interpunkcję**:
  - **Ignorowanie znaków specjalnych i interpunkcji**: Usuwa różnice wynikające z przecinków, kropek, myślników czy ukośników, dzięki czemu np. `"Ulica 5, Miasto"` jest równe `"Ulica 5 - Miasto"`.
  - **Ignorowanie nieregularnych spacji**: Redukuje wielokrotne spacje w polach scalonych.
  - **Pełne wsparcie dla polskich znaków**: Prawidłowo obsługuje polskie znaki diakrytyczne (`ą, ć, ę, ł, ń, ó, ś, ź, ż`) oraz liczby.
- **✨ Wyróżnianie zmienionych pól (komórek)**:
  - W wierszach o statusie **Zmieniony**, tylko faktycznie zmodyfikowane komórki są podświetlane kontrastowym tłem i pogrubioną czcionką.
- **🔄 Układ Góra / Dół (Pionowy) lub Obok siebie (Poziomy)**:
  - Domyślny układ **Góra / Dół (Pionowy)** gwarantujący wygodny podgląd szerokich tabel z wieloma kolumnami.
  - Możliwość natychmiastowej zmiany układu w oknie wyników.
- **🔀 Asymetryczne mapowanie kolumn (N:1 lub 1:N)**:
  - Możliwość łączenia wielu kolumn z pliku bazowego (np. `Imię` + `Nazwisko`) do porównania z jedną kolumną aktualizacji (`Imię i nazwisko`).
  - Tryby łączenia: `Concatenate` (z własnym separatorem), `Exact`, `FirstNonEmpty`.
- **🔑 Złożone klucze łączenia**:
  - Łączenie wierszy według jednego lub wielu pól kluczowych.
- **🌐 Trójjęzyczny interfejs**:
  - Język polski (PL), angielski (EN) oraz niemiecki (DE).
- **📊 Eksport wyników do Excela**:
  - Zapis raportu różnic do pliku `.xlsx` z oznaczeniem statusów i listy zmienionych pól.

---

## 💻 Wymagania wstępne

- **System operacyjny**: Windows 10 / Windows 11 / Windows Server 2016+.
- **PowerShell**: Windows PowerShell 5.1 lub PowerShell 7+.
- **Moduł**: `ImportExcel` (instalowany/ładowany automatycznie).

---

## 🚀 Uruchomienie

Wpisz w konsoli PowerShell:

```powershell
& 'D:\Skrypty\Mnich_Adam_Skrypty\Compare-ExcelFiles.ps1'
```

---

## 📖 Instrukcja krok po kroku

### Krok 1: Wybór plików i arkuszy (Zakładka 1: Pliki)
1. W sekcji **Plik bazowy** kliknij **"Wybierz plik bazowy"** i wskaż plik źródłowy (`.xlsx` lub `.csv`).
2. Wybierz odpowiedni arkusz z listy rozwijanej **Arkusz**.
3. W sekcji **Plik aktualizacji** kliknij **"Wybierz plik aktualizacji"** i wskaż plik z nowymi danymi.
4. Wybierz odpowiedni arkusz z listy rozwijanej **Arkusz**.

### Krok 2: Mapowanie kolumn (Zakładka 2: Mapowanie kolumn)
- **Automatycznie**: Kliknij **`⚡ Automatyczne mapowanie`**, aby utworzyć powiązania 1:1 dla identycznych nagłówków.
- **Ręcznie / Pola połączone**: Kliknij **"Dodaj regułę mapowania"**:
  - Zaznacz jedną lub więcej kolumn z pliku bazowego (Ctrl+Kliknięcie).
  - Zaznacz jedną lub więcej kolumn z pliku aktualizacji.
  - Wybierz tryb łączenia i separator (np. spacja lub ` | `).
  - Sprawdź podgląd w czasie rzeczywistym i zatwierdź **OK**.
- **Edycja / Usuwanie**: Zaznacz regułę w tabeli i wybierz **"Edytuj zaznaczone"** lub **"Usuń zaznaczone"**.

### Krok 3: Wybór klucza łączenia (Zakładka 3: Klucz łączenia)
- Zaznacz kolumnę lub zestaw kolumn jednoznacznie identyfikujących każdy rekord (np. `ID`, `PESEL`, `Numer`).

### Krok 4: Konfiguracja opcji (Zakładka 4: Opcje)
- **Układ widoku**: `Góra / Dół (Pionowy)` [Zalecany] lub `Obok siebie (Poziomy)`.
- **Motyw**: `☀️ Jasny` lub `🌙 Ciemny`.
- **Ignoruj wielkość liter przy porównywaniu**: Włączone domyślnie.
- **Przycinaj białe znaki przed porównaniem**: Włączone domyślnie.
- **Ignoruj znaki specjalne i interpunkcję**: Usuwa przecinki, myślniki, kropki podczas porównywania.
- **Ignoruj spacje i białe znaki przy porównywaniu pól**: Eliminuje różnice w liczbie spacji w polach scalonych.
- **Ukryj niezmienione wiersze w wynikach**: Pokazuje tylko wiersze dodane, usunięte i zmodyfikowane.

### Krok 5: Wykonanie porównania i analiza wyników
1. Kliknij **"▶ Uruchom porównanie"**.
2. W oknie wyników:
   - **Podsumowanie**: Liczba wierszy ogółem, Dodanych, Usuniętych, Zmienionych, Bez zmian.
   - **Filtrowanie**: Filtruj wg statusu lub wpisz frazę w pole wyszukiwania.
   - **Przełączanie motywu / układu**: Zmień motyw (`Ciemny`/`Jasny`) lub układ (`Pionowy`/`Poziomy`) w dowolnym momencie.
   - **Synchroniczne przewijanie**: Oba panele przewijają się jednocześnie.
   - **Eksport**: Kliknij **"Eksportuj"**, aby zapisać wynik do pliku `.xlsx`.

---
---

<a name="deutsche-bedienungsanleitung"></a>
# Deutsche Bedienungsanleitung

## 📌 Übersicht

**`Compare-ExcelFiles.ps1`** ist eine leistungsstarke PowerShell-WPF-Anwendung zum präzisen Vergleich zweier Excel- (`.xlsx`, `.xls`) oder `.csv`-Dateien mit unterschiedlichen Spaltenstrukturen, variierenden Spaltennamen oder kombinierten Datenfeldern.

### 🌟 Hauptfunktionen
- **🌙 Dynamischer Dunkel- (Dark Mode) und Hellmodus (Light Mode)**:
  - Vollständiger Dunkelmodus (`#0F172A` Schiefergrau, `#1E293B` Karten, `#334155` Steuerelemente, `#F8FAFC` kontrastreicher Text).
  - Native Integration der Windows 10/11 DWM-Titelleiste via `DwmSetWindowAttribute`.
  - Farbcodierte Zeilenstatus (Smaragdgrün für Hinzugefügt, Karminrot für Gelöscht, Dunkelbernstein für Geändert) und goldgelbe Hervorhebung geänderter Zellen (`#B45309` / `#FBBF24`).
  - Schnellumschaltung des Designs im Assistenten, in der Symbolleiste der Ergebnisansicht oder in den Optionen.
  - Automatische Speicherung der Design- und Spracheinstellungen in `%APPDATA%\Compare_ExcelFiles\settings.xml`.
- **⚡ Automatische Spaltenzuordnung (Auto-Map)**:
  - Automatische 1:1-Zuordnung aller übereinstimmenden Spaltennamen mit einem Klick.
  - Automatische Erkennung und Auswahl des Primärschlüssels (`ID`, `Lp`, `Code`, `Nr`, `Name`) in Registerkarte 3.
- **🔍 Intelligente Feldnormalisierung & Satzzeichen-Invarianz**:
  - **Ignorieren von Sonderzeichen und Satzzeichen**: Entfernt Kommas, Punkte, Bindestriche und Schrägstriche beim Vergleich.
  - **Ignorieren von Leerzeichendifferenzen**: Gleicht unregelmäßige Abstände in zusammengeführten Feldern aus.
  - **Volle Unicode-Unterstützung**: Erkennt deutsche Umlaute (`ä, ö, ü, ß`) und internationale Zeichen fehlerfrei.
- **✨ Hervorhebung geänderter Zellen**:
  - In geänderten Zeilen werden gezielt nur die tatsächlich modifizierten Zellen farblich hervorgehoben.
- **🔄 Duale Layout-Umschaltung (Oben/Unten vs. Nebeneinander)**:
  - Standardmäßig **Oben / Unten (Vertikal)** für breite Datentabellen mit vielen Spalten.
  - Jederzeitige Umschaltung auf **Nebeneinander (Horizontal)**.
- **🔀 Asymmetrische Spaltenzuordnung (N:1 oder 1:N)**:
  - Zusammenführen mehrerer Quellspalten (z. B. `Vorname` + `Nachname`) zum Vergleich mit einer Zielspalte (`Vollständiger Name`).
- **🔑 Zusammengesetzte Verknüpfungsschlüssel**:
  - Zeilenabgleich über einen oder mehrere zusammengesetzte Schlüssel.
- **🌐 Dreisprachige Benutzeroberfläche**:
  - Vollständige Unterstützung für Deutsch (DE), Englisch (EN) und Polnisch (PL).
- **📊 Formatierter Excel-Export**:
  - Speichern des Differenzberichts als formatierte `.xlsx`-Datei.

---

## 💻 Systemanforderungen

- **Betriebssystem**: Windows 10 / Windows 11 / Windows Server 2016+.
- **PowerShell**: Windows PowerShell 5.1 oder PowerShell Core 7+.
- **Modul**: `ImportExcel` (wird bei Bedarf automatisch installiert/geladen).

---

## 🚀 Starten der Anwendung

Führen Sie das Skript in PowerShell aus:

```powershell
& 'D:\Skrypty\Mnich_Adam_Skrypty\Compare-ExcelFiles.ps1'
```

---

## 📖 Schritt-für-Schritt-Anleitung

### Schritt 1: Dateien und Arbeitsblätter auswählen (Registerkarte 1: Dateien)
1. Klicken Sie auf **"Basisdatei auswählen"** und wählen Sie Ihre Referenzdatei.
2. Wählen Sie das gewünschte Arbeitsblatt im Dropdown-Menü **Arbeitsblatt**.
3. Klicken Sie auf **"Aktualisierungsdatei auswählen"** und wählen Sie Ihre neue Datei.
4. Wählen Sie das entsprechende Arbeitsblatt aus.

### Schritt 2: Spaltenzuordnung festlegen (Registerkarte 2: Spaltenzuordnung)
- **Automatisch**: Klicken Sie auf **`⚡ Automatische Zuordnung`**, um übereinstimmende Spalten 1:1 zu verknüpfen.
- **Manuell**: Klicken Sie auf **"Zuordnungsregel hinzufügen"** für kombinierte Spalten.
- **Zurücksetzen**: Klicken Sie auf **"Alles löschen"**, um die Liste zu leeren.

### Schritt 3: Verknüpfungsschlüssel auswählen (Registerkarte 3: Verknüpfungsschlüssel)
- Markieren Sie die Spalten, die jeden Datensatz eindeutig identifizieren (z. B. `ID`, `Personalnummer`).

### Schritt 4: Optionen konfigurieren (Registerkarte 4: Optionen)
- **Layout**: `Oben / Unten (Vertikal)` [Standard] oder `Nebeneinander (Horizontal)`.
- **Design**: `☀️ Hell` oder `🌙 Dunkel`.
- **Groß-/Kleinschreibung ignorieren**: Aktiviert.
- **Leerzeichen vor Vergleich kürzen**: Aktiviert.
- **Sonderzeichen und Satzzeichen ignorieren**: Aktiviert.
- **Alle internen Leerzeichenunterschiede ignorieren**: Aktiviert.
- **Unveränderte Zeilen ausblenden**: Deaktiviert.

### Schritt 5: Vergleich ausführen & auswerten
1. Klicken Sie auf **"▶ Vergleich starten"**.
2. Im Ergebnisfenster:
   - **Zusammenfassungsleiste**: Übersicht über Gesamtzahl, Hinzugefügte, Gelöschte, Geänderte und Unveränderte Zeilen.
   - **Filterleiste**: Nach Status filtern oder Volltextsuche nutzen.
   - **Synchrones Scrollen**: Beide Tabellen bewegen sich gleichzeitig.
   - **Exportieren**: Klicken Sie auf **"Exportieren"**, um den Bericht als Excel-Datei zu speichern.