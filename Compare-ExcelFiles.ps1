<#
.SYNOPSIS
    Compare-ExcelFiles - Excel File Comparison Tool with flexible column mapping.
.DESCRIPTION
    Compares a base Excel/CSV file against an update file. Supports asymmetric
    column structures (e.g. 3 base columns mapped to 1 update column), multi-column
    join keys, and shows results in a dual-grid WPF viewer with colour-coded rows.
    Added   = green, Deleted = red, Modified = yellow, Unchanged = white.
.AUTHOR
    Adam Mnich using Github Copilot
.NOTES
    2026.08.30 - Initial version
    - Requires ImportExcel module (will attempt to load/install if missing)
    - Can be compiled to EXE using ps2exe
#>



[CmdletBinding()]
param ()
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
#region 1. Dependencies and Modules
$ErrorActionPreferenceCurrent = $ErrorActionPreference
$ErrorActionPreference = 'Stop'

function Ensure-ImportExcelModule {
    try { Import-Module ImportExcel -ErrorAction Stop }
    catch {
        $imported = $false
        if ($env:USERDNSDOMAIN -match 'bgh.intra') {
            try {
                Import-Module '\\skatfs01\install$\scripts\ImportExcel\7.8.10\ImportExcel.psd1' -ErrorAction Stop
                $imported = $true
            }
            catch {}
        }
        if (-not $imported) {
            Write-Warning "Attempting to install 'ImportExcel' module..."
            try {
                Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber
                Import-Module ImportExcel -ErrorAction Stop
            }
            catch {
                Write-Error "The 'ImportExcel' module is required. Run: Install-Module ImportExcel -Scope CurrentUser"
                exit 1
            }
        }
    }
}
Ensure-ImportExcelModule

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$ErrorActionPreference = $ErrorActionPreferenceCurrent
#endregion

#region 2. Settings, Language and Theme

if (-not ('DwmHelper' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class DwmHelper {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction SilentlyContinue
}

if (-not ('FastDiffHelper' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

public static class FastDiffHelper {
    private static readonly Regex _rgxSpecial = new Regex(@"[^\p{L}\p{Nd}\s]", RegexOptions.Compiled);
    private static readonly Regex _rgxWhitespace = new Regex(@"\s+", RegexOptions.Compiled);

    public static string Normalize(string val, bool ignoreCase, bool trim, bool ignoreSpecial, bool ignoreAllWhitespace) {
        if (string.IsNullOrEmpty(val)) return string.Empty;
        string v = val;
        if (trim) v = v.Trim();
        if (ignoreCase) v = v.ToLowerInvariant();
        if (ignoreSpecial) v = _rgxSpecial.Replace(v, string.Empty);
        if (ignoreAllWhitespace) {
            v = _rgxWhitespace.Replace(v, string.Empty);
        } else {
            v = _rgxWhitespace.Replace(v, " ").Trim();
        }
        return v;
    }

    public static bool AreEqual(string bRaw, string uRaw, bool ignoreCase, bool trim, bool ignoreSpecial, bool ignoreAllSpaces) {
        if (string.Equals(bRaw, uRaw, StringComparison.Ordinal)) return true;
        if (bRaw == null) bRaw = string.Empty;
        if (uRaw == null) uRaw = string.Empty;
        if (bRaw.Length == 0 && uRaw.Length == 0) return true;

        if (ignoreCase && !ignoreSpecial && !ignoreAllSpaces && !trim) {
            return string.Equals(bRaw, uRaw, StringComparison.OrdinalIgnoreCase);
        }
        if (ignoreCase && !ignoreSpecial && !ignoreAllSpaces && trim) {
            return string.Equals(bRaw.Trim(), uRaw.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        string bNorm = Normalize(bRaw, ignoreCase, trim, ignoreSpecial, ignoreAllSpaces);
        string uNorm = Normalize(uRaw, ignoreCase, trim, ignoreSpecial, ignoreAllSpaces);
        return string.Equals(bNorm, uNorm, StringComparison.Ordinal);
    }

    public static string MergeValues(IList<string> values, string mergeMode, string separator, bool trim) {
        if (values == null || values.Count == 0) return string.Empty;
        if (string.Equals(mergeMode, "Exact", StringComparison.OrdinalIgnoreCase)) {
            string v = values[0] ?? string.Empty;
            return trim ? v.Trim() : v;
        }
        if (string.Equals(mergeMode, "FirstNonEmpty", StringComparison.OrdinalIgnoreCase)) {
            for (int i = 0; i < values.Count; i++) {
                string v = values[i];
                if (!string.IsNullOrEmpty(v)) {
                    if (trim) v = v.Trim();
                    if (!string.IsNullOrEmpty(v)) return v;
                }
            }
            return string.Empty;
        }
        StringBuilder sb = new StringBuilder();
        string sep = separator ?? string.Empty;
        for (int i = 0; i < values.Count; i++) {
            string v = values[i] ?? string.Empty;
            if (trim) v = v.Trim();
            if (i > 0) sb.Append(sep);
            sb.Append(v);
        }
        return sb.ToString();
    }
}
"@ -ErrorAction SilentlyContinue
}

function Set-WindowDarkMode {
    param(
        [System.Windows.Window]$Window = $null,
        [IntPtr]$WindowHandle = [IntPtr]::Zero,
        [bool]$IsDark = $false
    )
    try {
        $hwnd = $WindowHandle
        if ($hwnd -eq [IntPtr]::Zero -and $null -ne $Window) {
            $helper = New-Object System.Windows.Interop.WindowInteropHelper($Window)
            $hwnd = $helper.Handle
        }
        if ($hwnd -ne [IntPtr]::Zero) {
            $val = if ($IsDark) { 1 } else { 0 }
            [DwmHelper]::DwmSetWindowAttribute($hwnd, 20, [ref]$val, 4) | Out-Null
            [DwmHelper]::DwmSetWindowAttribute($hwnd, 19, [ref]$val, 4) | Out-Null
        }
    }
    catch {}
}

$Global:CurrentLang = 'PL'
$Global:CurrentTheme = 'Light'
$script:SettingsPath = Join-Path $env:APPDATA 'Compare_ExcelFiles'
$script:SettingsFile = Join-Path $script:SettingsPath 'settings.xml'
if (-not (Test-Path $script:SettingsPath)) { New-Item -ItemType Directory -Path $script:SettingsPath -Force | Out-Null }

function Save-AppSettings {
    param([string]$Lang = $Global:CurrentLang, [string]$Theme = $Global:CurrentTheme)
    ([pscustomobject]@{ Language = $Lang; Theme = $Theme }) | Export-Clixml -Path $script:SettingsFile -Force
}

function Load-AppSettings {
    if (Test-Path $script:SettingsFile) {
        try {
            $s = Import-Clixml $script:SettingsFile
            if ($s.Language -in 'EN', 'PL', 'DE') { $Global:CurrentLang = $s.Language }
            if ($s.Theme -in 'Light', 'Dark') { $Global:CurrentTheme = $s.Theme }
        }
        catch {}
    }
}
try { Load-AppSettings } catch { $Global:CurrentLang = 'PL'; $Global:CurrentTheme = 'Light' }

function Get-ThemePalette {
    param([string]$Theme = $Global:CurrentTheme)
    if ($Theme -eq 'Dark') {
        return @{
            IsDark          = $true
            BgWindow        = '#0F172A'
            BgHeader        = '#0F172A'
            BgCard          = '#1E293B'
            BgInput         = '#334155'
            BgButtonDefault = '#334155'
            BgButtonHover   = '#475569'
            TextPrimary     = '#F8FAFC'
            TextSecondary   = '#CBD5E1'
            TextMuted       = '#94A3B8'
            Border          = '#475569'
            BorderSubtle    = '#334155'
            GridLines       = '#334155'
            RowBg           = '#1E293B'
            RowAltBg        = '#0F172A'
            HeaderBg        = '#334155'
            HeaderFg        = '#F8FAFC'
            SummaryBg       = '#020617'
            RowAddedBg      = '#064E3B'
            RowAddedFg      = '#ECFDF5'
            RowDeletedBg    = '#7F1D1D'
            RowDeletedFg    = '#FEF2F2'
            RowModifiedBg   = '#713F12'
            RowModifiedFg   = '#FEFCE8'
            RowUnchangedBg  = '#1E293B'
            RowUnchangedFg  = '#F8FAFC'
            CellChgBg       = '#B45309'
            CellChgBorder   = '#FBBF24'
            CellChgFg       = '#FEF08A'
            TabActiveBg     = '#1E293B'
            TabInactiveBg   = '#0F172A'
            TabHoverBg      = '#334155'
            ItemHoverBg     = '#334155'
        }
    }
    else {
        return @{
            IsDark          = $false
            BgWindow        = '#F3F4F6'
            BgHeader        = '#1F2937'
            BgCard          = '#FFFFFF'
            BgInput         = '#FFFFFF'
            BgButtonDefault = '#E5E7EB'
            BgButtonHover   = '#D1D5DB'
            TextPrimary     = '#1F2937'
            TextSecondary   = '#4B5563'
            TextMuted       = '#6B7280'
            Border          = '#D1D5DB'
            BorderSubtle    = '#E5E7EB'
            GridLines       = '#E5E7EB'
            RowBg           = '#FFFFFF'
            RowAltBg        = '#F9FAFB'
            HeaderBg        = '#F3F4F6'
            HeaderFg        = '#1F2937'
            SummaryBg       = '#111827'
            RowAddedBg      = '#DCFCE7'
            RowAddedFg      = '#1F2937'
            RowDeletedBg    = '#FEE2E2'
            RowDeletedFg    = '#1F2937'
            RowModifiedBg   = '#FEF9C3'
            RowModifiedFg   = '#1F2937'
            RowUnchangedBg  = '#FFFFFF'
            RowUnchangedFg  = '#1F2937'
            CellChgBg       = '#FDE047'
            CellChgBorder   = '#D97706'
            CellChgFg       = '#78350F'
            TabActiveBg     = '#FFFFFF'
            TabInactiveBg   = '#E5E7EB'
            TabHoverBg      = '#D1D5DB'
            ItemHoverBg     = '#E5E7EB'
        }
    }
}

function Update-WpfThemeResources {
    param([System.Windows.Window]$Window, [hashtable]$Palette)
    $bc = New-Object System.Windows.Media.BrushConverter
    foreach ($key in $Palette.Keys) {
        $val = $Palette[$key]
        if ($val -is [string] -and $val.StartsWith('#')) {
            $brushName = "${key}Brush"
            $brush = $bc.ConvertFromString($val)
            if ($brush.CanFreeze) { $brush.Freeze() }
            $Window.Resources[$brushName] = $brush
        }
    }
}

$Global:Translations = @{
    EN = @{
        WizardTitle           = 'Excel File Comparator - Setup'
        Language              = 'Language:'
        ThemeLabel            = 'Theme:'
        ThemeLight            = '☀️ Light'
        ThemeDark             = '🌙 Dark'
        BtnThemeToggle        = 'Theme'
        TabFiles              = 'Files'
        TabMapping            = 'Column Mapping'
        TabJoinKey            = 'Join Key'
        TabOptions            = 'Options'
        BtnBaseFile           = 'Select Base File'
        BtnUpdateFile         = 'Select Update File'
        SheetLabel            = 'Sheet:'
        LblBaseFile           = 'Base File (reference):'
        LblUpdateFile         = 'Update File (changes):'
        BtnAddMapping         = 'Add Mapping Rule'
        BtnEditMapping        = 'Edit Selected'
        BtnRemoveMapping      = 'Remove Selected'
        BtnAutoMap            = '⚡ Auto-Map Columns'
        BtnClearMap           = 'Clear All'
        MsgAutoMapped         = 'Successfully auto-mapped {0} matching column(s).'
        WarnNoMatchingCols    = 'No columns with matching names were found between the two files.'
        ColBaseColumns        = 'Base Column(s)'
        ColMergeMode          = 'Merge Mode'
        ColSeparator          = 'Separator'
        ColUpdateColumns      = 'Update Column(s)'
        ColLabel              = 'Label'
        LblJoinKey            = 'Select column(s) used to match rows between files:'
        LblJoinBase           = 'Base file join column(s):'
        LblJoinUpdate         = 'Update file join column(s):'
        LblIgnoreCase         = 'Ignore case when comparing'
        LblTrimWhitespace     = 'Trim whitespace before comparing'
        LblIgnoreSpecialChars = 'Ignore punctuation & special characters (e.g. , . - / \)'
        LblIgnoreAllSpaces    = 'Ignore internal whitespace differences when comparing fields'
        LblIgnoreUnchanged    = 'Hide Unchanged rows in results'
        BtnRunCompare         = 'Run Comparison'
        StatusAdded           = 'Added'
        StatusDeleted         = 'Deleted'
        StatusModified        = 'Modified'
        StatusUnchanged       = 'Unchanged'
        StatusAll             = 'All'
        DashSearch            = 'Search (all fields):'
        DashReset             = 'Reset Filters'
        DashExport            = 'Export to Excel'
        ColDiffStatus         = 'Status'
        ColChangedCols        = 'Changed Columns'
        PanelBase             = 'Base File'
        PanelUpdate           = 'Update File'
        LoadingData           = 'Running comparison, please wait...'
        WarnNoBase            = 'Please select the Base file first.'
        WarnNoUpdate          = 'Please select the Update file first.'
        WarnNoMapping         = 'Please add at least one column mapping rule.'
        WarnNoJoinKey         = 'Please select at least one join key column.'
        WarnJoinMismatch      = 'Number of join key columns must match between base and update.'
        MapDialogTitle        = 'Column Mapping Rule'
        MapBaseLabel          = 'Base column(s): (hold Ctrl for multi-select)'
        MapUpdateLabel        = 'Update column(s): (hold Ctrl for multi-select)'
        MapMergeLabel         = 'Merge mode:'
        MapSepLabel           = 'Separator (for Concatenate):'
        MapLabelLabel         = 'Display label:'
        MapOK                 = 'OK'
        MapCancel             = 'Cancel'
        ExportSaved           = 'Diff exported to:'
        SummaryAdded          = 'Added'
        SummaryDeleted        = 'Deleted'
        SummaryModified       = 'Modified'
        SummaryUnchanged      = 'Unchanged'
        SummaryTotal          = 'Total'
        DialogWarn            = 'Warning'
        DialogInfo            = 'Information'
        SyncScroll            = 'Sync Scroll'
        LayoutMode            = 'Layout:'
        LayoutVert            = 'Top / Bottom (Vertical)'
        LayoutHoriz           = 'Side by Side (Horizontal)'
        BtnClearRowFilter     = '✕ Clear Row Filter'
        TipRowFilter          = 'Click to filter both files to this joined row (click again or press Esc to show all)'
        JoinModeKey           = 'Match rows using Key Column(s)'
        JoinModeRow           = 'Match rows by Row Order (Sequential / Line-by-Line)'
        TipMatchByRow         = 'Compares Row 1 to Row 1, Row 2 to Row 2 sequentially. No key columns required.'
        LblRowIndexJoin       = 'Sequential row comparison active: rows are matched line-by-line (Row 1 vs Row 1, Row 2 vs Row 2...)'
        RowPrefix             = 'Row'
    }
    PL = @{
        WizardTitle           = 'Porównywarka plików Excel - Konfiguracja'
        Language              = 'Język:'
        ThemeLabel            = 'Motyw:'
        ThemeLight            = '☀️ Jasny'
        ThemeDark             = '🌙 Ciemny'
        BtnThemeToggle        = 'Motyw'
        TabFiles              = 'Pliki'
        TabMapping            = 'Mapowanie kolumn'
        TabJoinKey            = 'Klucz łączenia'
        TabOptions            = 'Opcje'
        BtnBaseFile           = 'Wybierz plik bazowy'
        BtnUpdateFile         = 'Wybierz plik aktualizacji'
        SheetLabel            = 'Arkusz:'
        LblBaseFile           = 'Plik bazowy (referencyjny):'
        LblUpdateFile         = 'Plik aktualizacji (zmiany):'
        BtnAddMapping         = 'Dodaj regułę mapowania'
        BtnEditMapping        = 'Edytuj zaznaczone'
        BtnRemoveMapping      = 'Usuń zaznaczone'
        BtnAutoMap            = '⚡ Automatyczne mapowanie'
        BtnClearMap           = 'Wyczyść wszystko'
        MsgAutoMapped         = 'Pomyślnie zmapowano automatycznie {0} pasujących kolumn.'
        WarnNoMatchingCols    = 'Nie znaleziono kolumn o pasujących nazwach między plikami.'
        ColBaseColumns        = 'Kolumna(y) bazowe'
        ColMergeMode          = 'Tryb łączenia'
        ColSeparator          = 'Separator'
        ColUpdateColumns      = 'Kolumna(y) aktualizacji'
        ColLabel              = 'Etykieta'
        LblJoinKey            = 'Wybierz kolumny identyfikujące wiersze między plikami:'
        LblJoinBase           = 'Kolumny klucza w pliku bazowym:'
        LblJoinUpdate         = 'Kolumny klucza w pliku aktualizacji:'
        LblIgnoreCase         = 'Ignoruj wielkość liter przy porównywaniu'
        LblTrimWhitespace     = 'Przycinaj białe znaki przed porównaniem'
        LblIgnoreSpecialChars = 'Ignoruj znaki specjalne i interpunkcję (np. przecinki, myślniki)'
        LblIgnoreAllSpaces    = 'Ignoruj spacje i białe znaki przy porównywaniu pól'
        LblIgnoreUnchanged    = 'Ukryj niezmienione wiersze w wynikach'
        BtnRunCompare         = 'Uruchom porównanie'
        StatusAdded           = 'Dodany'
        StatusDeleted         = 'Usunięty'
        StatusModified        = 'Zmieniony'
        StatusUnchanged       = 'Bez zmian'
        StatusAll             = 'Wszystkie'
        DashSearch            = 'Szukaj (wszystkie pola):'
        DashReset             = 'Resetuj filtry'
        DashExport            = 'Eksportuj do Excel'
        ColDiffStatus         = 'Status'
        ColChangedCols        = 'Zmienione kolumny'
        PanelBase             = 'Plik bazowy'
        PanelUpdate           = 'Plik aktualizacji'
        LoadingData           = 'Trwa porównywanie, proszę czekać...'
        WarnNoBase            = 'Proszę wybrać plik bazowy.'
        WarnNoUpdate          = 'Proszę wybrać plik aktualizacji.'
        WarnNoMapping         = 'Proszę dodać co najmniej jedną regułę mapowania kolumn.'
        WarnNoJoinKey         = 'Proszę wybrać co najmniej jedną kolumnę klucza łączenia.'
        WarnJoinMismatch      = 'Liczba kolumn klucza musi być taka sama dla obu plików.'
        MapDialogTitle        = 'Reguła mapowania kolumn'
        MapBaseLabel          = 'Kolumny bazowe: (przytrzymaj Ctrl dla wielu)'
        MapUpdateLabel        = 'Kolumny aktualizacji: (przytrzymaj Ctrl dla wielu)'
        MapMergeLabel         = 'Tryb łączenia:'
        MapSepLabel           = 'Separator (dla Concatenate):'
        MapLabelLabel         = 'Etykieta wyświetlania:'
        MapOK                 = 'OK'
        MapCancel             = 'Anuluj'
        ExportSaved           = 'Diff zapisany do:'
        SummaryAdded          = 'Dodane'
        SummaryDeleted        = 'Usunięte'
        SummaryModified       = 'Zmienione'
        SummaryUnchanged      = 'Bez zmian'
        SummaryTotal          = 'Łącznie'
        DialogWarn            = 'Ostrzeżenie'
        DialogInfo            = 'Informacja'
        SyncScroll            = 'Synchronizuj przewijanie'
        LayoutMode            = 'Układ widoku:'
        LayoutVert            = 'Góra / Dół (Pionowy)'
        LayoutHoriz           = 'Obok siebie (Poziomy)'
        BtnClearRowFilter     = '✕ Wyczyść filtr wiersza'
        TipRowFilter          = 'Kliknij, aby przefiltrować oba pliki do tego złączonego wiersza (ponowne kliknięcie lub Esc przywraca wszystkie)'
        JoinModeKey           = 'Łącz wiersze używając kolumn klucza'
        JoinModeRow           = 'Łącz wiersze według kolejności (wiersz po wierszu)'
        TipMatchByRow         = 'Porównuje wiersz 1 z 1, wiersz 2 z 2 sekwencyjnie. Nie wymaga kolumn klucza.'
        LblRowIndexJoin       = 'Aktywne porównywanie sekwencyjne: wiersze są łączone po kolei (Wiersz 1 vs Wiersz 1, Wiersz 2 vs Wiersz 2...)'
        RowPrefix             = 'Wiersz'
    }
    DE = @{
        WizardTitle           = 'Excel-Vergleich - Konfiguration'
        Language              = 'Sprache:'
        ThemeLabel            = 'Design:'
        ThemeLight            = '☀️ Hell'
        ThemeDark             = '🌙 Dunkel'
        BtnThemeToggle        = 'Design'
        TabFiles              = 'Dateien'
        TabMapping            = 'Spaltenzuordnung'
        TabJoinKey            = 'Verknuepfungsschluessel'
        TabOptions            = 'Optionen'
        BtnBaseFile           = 'Basisdatei auswaehlen'
        BtnUpdateFile         = 'Aktualisierungsdatei auswaehlen'
        SheetLabel            = 'Arbeitsblatt:'
        LblBaseFile           = 'Basisdatei (Referenz):'
        LblUpdateFile         = 'Aktualisierungsdatei (Aenderungen):'
        BtnAddMapping         = 'Zuordnungsregel hinzufuegen'
        BtnEditMapping        = 'Ausgewaehlte bearbeiten'
        BtnRemoveMapping      = 'Ausgewaehlte entfernen'
        BtnAutoMap            = '⚡ Automatisch zuordnen'
        BtnClearMap           = 'Alle leeren'
        MsgAutoMapped         = '{0} uebereinstimmende Spalte(n) wurden automatisch zugeordnet.'
        WarnNoMatchingCols    = 'Keine uebereinstimmenden Spaltennamen zwischen den Dateien gefunden.'
        ColBaseColumns        = 'Basisspalte(n)'
        ColMergeMode          = 'Zusammenfuehrungsmodus'
        ColSeparator          = 'Trennzeichen'
        ColUpdateColumns      = 'Aktualisierungsspalte(n)'
        ColLabel              = 'Bezeichnung'
        LblJoinKey            = 'Spalten zur Zeilenidentifikation zwischen Dateien:'
        LblJoinBase           = 'Schlusselspalten Basisdatei:'
        LblJoinUpdate         = 'Schlusselspalten Aktualisierungsdatei:'
        LblIgnoreCase         = 'Gross-/Kleinschreibung ignorieren'
        LblTrimWhitespace     = 'Leerzeichen beim Vergleich entfernen'
        LblIgnoreSpecialChars = 'Satz- und Sonderzeichen ignorieren (z. B. , . - / \)'
        LblIgnoreAllSpaces    = 'Leerzeichen beim Feldvergleich ignorieren'
        LblIgnoreUnchanged    = 'Unveraenderte Zeilen in Ergebnissen ausblenden'
        BtnRunCompare         = 'Vergleich starten'
        StatusAdded           = 'Hinzugefuegt'
        StatusDeleted         = 'Geloescht'
        StatusModified        = 'Geaendert'
        StatusUnchanged       = 'Unveraendert'
        StatusAll             = 'Alle'
        DashSearch            = 'Suche (alle Felder):'
        DashReset             = 'Filter zuruecksetzen'
        DashExport            = 'Nach Excel exportieren'
        ColDiffStatus         = 'Status'
        ColChangedCols        = 'Geaenderte Spalten'
        PanelBase             = 'Basisdatei'
        PanelUpdate           = 'Aktualisierungsdatei'
        LoadingData           = 'Vergleich wird durchgefuehrt, bitte warten...'
        WarnNoBase            = 'Bitte waehlen Sie die Basisdatei aus.'
        WarnNoUpdate          = 'Bitte waehlen Sie die Aktualisierungsdatei aus.'
        WarnNoMapping         = 'Bitte fuegen Sie mindestens eine Spaltenzuordnungsregel hinzu.'
        WarnNoJoinKey         = 'Bitte waehlen Sie mindestens eine Schlusselspalte.'
        WarnJoinMismatch      = 'Die Anzahl der Schlusselspalten muss fuer beide Dateien uebereinstimmen.'
        MapDialogTitle        = 'Spaltenzuordnungsregel'
        MapBaseLabel          = 'Basisspalten: (Strg fuer Mehrfachauswahl)'
        MapUpdateLabel        = 'Aktualisierungsspalten: (Strg fuer Mehrfachauswahl)'
        MapMergeLabel         = 'Zusammenfuehrungsmodus:'
        MapSepLabel           = 'Trennzeichen (fuer Zusammenfuehren):'
        MapLabelLabel         = 'Anzeigebezeichnung:'
        MapOK                 = 'OK'
        MapCancel             = 'Abbrechen'
        ExportSaved           = 'Diff exportiert nach:'
        SummaryAdded          = 'Hinzugefuegt'
        SummaryDeleted        = 'Geloescht'
        SummaryModified       = 'Geaendert'
        SummaryUnchanged      = 'Unveraendert'
        SummaryTotal          = 'Gesamt'
        DialogWarn            = 'Warnung'
        DialogInfo            = 'Information'
        SyncScroll            = 'Scroll synchronisieren'
        LayoutMode            = 'Duales Layout:'
        LayoutVert            = 'Oben / Unten (Vertikal)'
        LayoutHoriz           = 'Nebeneinander (Horizontal)'
        BtnClearRowFilter     = '✕ Zeilenfilter aufheben'
        TipRowFilter          = 'Klicken, um beide Dateien auf diese verknuepfte Zeile zu filtern (erneuter Klick oder Esc zeigt alle)'
        JoinModeKey           = 'Zeilen anhand von Schlusselspalten verknuepfen'
        JoinModeRow           = 'Zeilen nach Zeilenreihenfolge verknuepfen (Zeile fuer Zeile)'
        TipMatchByRow         = 'Vergleicht Zeile 1 mit Zeile 1, Zeile 2 mit Zeile 2 sequenziell. Keine Schlusselspalten erforderlich.'
        LblRowIndexJoin       = 'Sequenzieller Zeilenvergleich aktiv: Zeilen werden der Reihe nach verknuepft (Zeile 1 vs Zeile 1, Zeile 2 vs Zeile 2...)'
        RowPrefix             = 'Zeile'
    }
}

function Get-Loc {
    param([string]$Key)
    $val = $Global:Translations[$Global:CurrentLang][$Key]
    if ($null -eq $val) { $val = $Global:Translations['EN'][$Key] }
    if ($null -eq $val) { $val = $Key }
    return $val
}
#endregion

#region 3. Helper Functions

function Get-SafePropName ([string]$ColName) {
    return "_IsChanged_" + ($ColName -replace '[^a-zA-Z0-9_]', '_')
}

function Get-ScrollViewer {
    param ($Control)
    if ($null -eq $Control) { return $null }
    if ($Control -is [System.Windows.Controls.ScrollViewer]) { return $Control }
    $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Control)
    for ($i = 0; $i -lt $count; $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Control, $i)
        $result = Get-ScrollViewer $child
        if ($null -ne $result) { return $result }
    }
    return $null
}

function Get-VisualParentRow ([System.Windows.DependencyObject]$depObj) {
    while ($null -ne $depObj) {
        if ($depObj -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or
            $depObj -is [System.Windows.Controls.Primitives.ScrollBar] -or
            $depObj -is [System.Windows.Controls.Primitives.Thumb]) {
            return $null
        }
        if ($depObj -is [System.Windows.Controls.DataGridRow]) {
            return $depObj
        }
        $depObj = [System.Windows.Media.VisualTreeHelper]::GetParent($depObj)
    }
    return $null
}

function Pick-File ([System.Windows.Controls.TextBlock]$pb, [System.Windows.Controls.ComboBox]$sc) {
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = 'Excel & CSV Files|*.xlsx;*.xls;*.csv|Excel Files|*.xlsx;*.xls|CSV Files|*.csv'
    if ($fd.ShowDialog() -eq 'OK') {
        $pb.Text = $fd.FileName; $sc.Items.Clear()
        if ($fd.FileName -match '\.csv$') {
            $sc.IsEnabled = $false; [void]$sc.Items.Add('(CSV)'); $sc.SelectedIndex = 0
        }
        else {
            $sc.IsEnabled = $true
            $sheets = (Get-ExcelSheetInfo -Path $fd.FileName).Name
            foreach ($s in $sheets) { [void]$sc.Items.Add($s) }
            if ($sheets.Count -gt 0) { $sc.SelectedIndex = 0 }
        }
        return $fd.FileName
    }
    return $null
}

function Get-ExcelHeaders {
    param ([string]$Path, [string]$Sheet)
    try {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        if ($Path -match '\.csv$') {
            $header = Get-Content -Path $Path -Encoding UTF8 |
            Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -First 1
            $delims = @{
                ','  = ($header.ToCharArray() | Where-Object { $_ -eq ',' }).Count
                ';'  = ($header.ToCharArray() | Where-Object { $_ -eq ';' }).Count
                "`t" = ($header.ToCharArray() | Where-Object { $_ -eq "`t" }).Count
            }
            $script:Delimiter = ($delims.GetEnumerator() |
                Where-Object { $_.Value -gt 0 } | Sort-Object Value -Descending | Select-Object -First 1).Key
            if (-not $script:Delimiter) { $script:Delimiter = (Get-Culture).TextInfo.ListSeparator }
            $data = Import-Csv -Path $Path -Delimiter $script:Delimiter | Select-Object -First 1
            if ($data) { return $data.psobject.properties.name }
        }
        else {
            $data = Import-Excel -Path $Path -WorksheetName $Sheet -EndRow 2 -ErrorAction Stop
            if ($data) {
                return $data[0].psobject.properties.name |
                Where-Object { $_ -notmatch '^(RowError|RowState|Table|ItemArray|HasErrors)$' }
            }
        }
    }
    catch { }
    finally { [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default }
    return @()
}

function Load-ExcelData {
    param ([string]$Path, [string]$Sheet)
    if ($Path -match '\.csv$') { return @(Import-Csv -Path $Path -Delimiter $script:Delimiter) }
    elseif ($Sheet -and $Sheet -ne '(CSV)') { return @(Import-Excel -Path $Path -WorksheetName $Sheet) }
    else { return @(Import-Excel -Path $Path) }
}

function Normalize-CompareValue {
    param (
        [string]$Value,
        [bool]$IgnoreCase = $true,
        [bool]$TrimWhitespace = $true,
        [bool]$IgnoreSpecialChars = $false,
        [bool]$IgnoreAllWhitespace = $false
    )
    return [FastDiffHelper]::Normalize($Value, $IgnoreCase, $TrimWhitespace, $IgnoreSpecialChars, $IgnoreAllWhitespace)
}

function Get-MergedValue {
    param (
        [object]$Row,
        [string[]]$Columns,
        [string]$MergeMode,
        [string]$Separator,
        [bool]$Trim = $true,
        [bool]$IgnoreCase = $true,
        [bool]$IgnoreSpecialChars = $false,
        [bool]$IgnoreAllWhitespace = $false
    )
    if ($MergeMode -eq 'Exact' -and $Columns.Length -eq 1) {
        $raw = if ($null -ne $Row.($Columns[0])) { $Row.($Columns[0]).ToString() } else { '' }
        return [FastDiffHelper]::Normalize($raw, $IgnoreCase, $Trim, $IgnoreSpecialChars, $IgnoreAllWhitespace)
    }
    $vals = [System.Collections.Generic.List[string]]::new($Columns.Length)
    foreach ($col in $Columns) {
        $vals.Add((if ($null -ne $Row.$col) { $Row.$col.ToString() } else { '' }))
    }
    $merged = [FastDiffHelper]::MergeValues($vals, $MergeMode, $Separator, $Trim)
    return [FastDiffHelper]::Normalize($merged, $IgnoreCase, $Trim, $IgnoreSpecialChars, $IgnoreAllWhitespace)
}

function Build-JoinKey {
    param (
        [object]$Row,
        [string[]]$KeyColumns,
        [bool]$Trim = $true,
        [bool]$IgnoreCase = $true,
        [bool]$IgnoreSpecialChars = $false,
        [bool]$IgnoreAllWhitespace = $false
    )
    $sb = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $KeyColumns.Length; $i++) {
        $col = $KeyColumns[$i]
        $v = if ($null -ne $Row.$col) { $Row.$col.ToString() } else { '' }
        $norm = [FastDiffHelper]::Normalize($v, $IgnoreCase, $Trim, $IgnoreSpecialChars, $IgnoreAllWhitespace)
        if ($i -gt 0) { [void]$sb.Append('|||') }
        [void]$sb.Append($norm)
    }
    return $sb.ToString()
}

function Invoke-ExcelDiff {
    param (
        [object[]]$BaseData,
        [object[]]$UpdateData,
        [string[]]$BaseAllProps,
        [string[]]$UpdateAllProps,
        [hashtable[]]$MappingRules,
        [string[]]$JoinBaseColumns,
        [string[]]$JoinUpdateColumns,
        [bool]$IgnoreCase,
        [bool]$TrimWhitespace,
        [bool]$IgnoreSpecialChars = $true,
        [bool]$IgnoreAllSpaces = $true,
        [bool]$MatchByRowOrder = $false
    )

    $statusAdded = Get-Loc 'StatusAdded'
    $statusDeleted = Get-Loc 'StatusDeleted'
    $statusModified = Get-Loc 'StatusModified'
    $statusUnchanged = Get-Loc 'StatusUnchanged'

    # Precompute property safe names to avoid repeated regex in tight loops
    $baseSafeProps = @{}
    foreach ($col in $BaseAllProps) { $baseSafeProps[$col] = Get-SafePropName $col }
    $updateSafeProps = @{}
    foreach ($col in $UpdateAllProps) { $updateSafeProps[$col] = Get-SafePropName $col }

    # Precompile mapping rules for direct 1-to-1 fast-path
    $compiledRules = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $MappingRules) {
        $isDirect = ($r.MergeMode -eq 'Exact' -and $r.BaseColumns.Length -eq 1 -and $r.UpdateColumns.Length -eq 1)
        $compiledRules.Add([PSCustomObject]@{
            IsDirect      = $isDirect
            BaseCol       = if ($isDirect) { $r.BaseColumns[0] } else { $null }
            UpdateCol     = if ($isDirect) { $r.UpdateColumns[0] } else { $null }
            BaseColumns   = $r.BaseColumns
            UpdateColumns = $r.UpdateColumns
            MergeMode     = $r.MergeMode
            Separator     = $r.Separator
            Label         = $r.Label
        })
    }

    # Region: Sequential Row-by-Row Comparison Mode
    if ($MatchByRowOrder) {
        $rowPrefix = Get-Loc 'RowPrefix'
        if ([string]::IsNullOrWhiteSpace($rowPrefix)) { $rowPrefix = 'Row' }

        $resultBase = [System.Collections.ArrayList]::new()
        $resultUpdate = [System.Collections.ArrayList]::new()

        $bTotal = if ($null -ne $BaseData) { $BaseData.Count } else { 0 }
        $uTotal = if ($null -ne $UpdateData) { $UpdateData.Count } else { 0 }
        $maxRows = [Math]::Max($bTotal, $uTotal)

        for ($i = 0; $i -lt $maxRows; $i++) {
            $bRow = if ($i -lt $bTotal) { $BaseData[$i] } else { $null }
            $uRow = if ($i -lt $uTotal) { $UpdateData[$i] } else { $null }

            $pairIdx = $resultBase.Count
            $pairId = "P_$pairIdx"
            $key = "$rowPrefix $($i + 1)"

            if ($null -eq $bRow) {
                # Added - only in update
                $sbUpd = [System.Text.StringBuilder]::new()
                $uProps = [ordered]@{
                    _DiffStatus     = $statusAdded
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $UpdateAllProps) {
                    $v = $uRow.$p
                    $uProps[$p] = $v
                    $uProps[$updateSafeProps[$p]] = $false
                    if ($null -ne $v) { [void]$sbUpd.Append($v.ToString()); [void]$sbUpd.Append('|') }
                }
                $uProps['_SearchableText'] = $sbUpd.ToString()

                $blank = [ordered]@{
                    _DiffStatus     = $statusAdded
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $BaseAllProps) {
                    $blank[$p] = ''
                    $blank[$baseSafeProps[$p]] = $false
                }
                $blank['_SearchableText'] = ''

                [void]$resultBase.Add([PSCustomObject]$blank)
                [void]$resultUpdate.Add([PSCustomObject]$uProps)
            }
            elseif ($null -eq $uRow) {
                # Deleted - only in base
                $sbBase = [System.Text.StringBuilder]::new()
                $bProps = [ordered]@{
                    _DiffStatus     = $statusDeleted
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $BaseAllProps) {
                    $v = $bRow.$p
                    $bProps[$p] = $v
                    $bProps[$baseSafeProps[$p]] = $false
                    if ($null -ne $v) { [void]$sbBase.Append($v.ToString()); [void]$sbBase.Append('|') }
                }
                $bProps['_SearchableText'] = $sbBase.ToString()

                $blank = [ordered]@{
                    _DiffStatus     = $statusDeleted
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $UpdateAllProps) {
                    $blank[$p] = ''
                    $blank[$updateSafeProps[$p]] = $false
                }
                $blank['_SearchableText'] = ''

                [void]$resultBase.Add([PSCustomObject]$bProps)
                [void]$resultUpdate.Add([PSCustomObject]$blank)
            }
            else {
                # Both exist - compare using mapping rules
                $changedLabels = [System.Collections.Generic.List[string]]::new()
                $changedBaseCols = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $changedUpdateCols = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

                foreach ($rule in $compiledRules) {
                    $isMatch = $false
                    if ($rule.IsDirect) {
                        $bValObj = $bRow.($rule.BaseCol)
                        $bRaw = if ($null -ne $bValObj) { $bValObj.ToString() } else { '' }
                        $uValObj = $uRow.($rule.UpdateCol)
                        $uRaw = if ($null -ne $uValObj) { $uValObj.ToString() } else { '' }
                        $isMatch = [FastDiffHelper]::AreEqual($bRaw, $uRaw, $IgnoreCase, $TrimWhitespace, $IgnoreSpecialChars, $IgnoreAllSpaces)
                    }
                    else {
                        $bVal = Get-MergedValue -Row $bRow -Columns $rule.BaseColumns   -MergeMode $rule.MergeMode -Separator $rule.Separator -Trim $TrimWhitespace -IgnoreCase $IgnoreCase -IgnoreSpecialChars $IgnoreSpecialChars -IgnoreAllWhitespace $IgnoreAllSpaces
                        $uVal = Get-MergedValue -Row $uRow -Columns $rule.UpdateColumns -MergeMode $rule.MergeMode -Separator $rule.Separator -Trim $TrimWhitespace -IgnoreCase $IgnoreCase -IgnoreSpecialChars $IgnoreSpecialChars -IgnoreAllWhitespace $IgnoreAllSpaces
                        $isMatch = ($bVal -eq $uVal)
                    }

                    if (-not $isMatch) {
                        $changedLabels.Add($rule.Label)
                        foreach ($c in $rule.BaseColumns) { [void]$changedBaseCols.Add($c) }
                        foreach ($c in $rule.UpdateColumns) { [void]$changedUpdateCols.Add($c) }
                    }
                }

                $status = if ($changedLabels.Count -gt 0) { $statusModified } else { $statusUnchanged }
                $changed = $changedLabels -join ', '

                $sbBase = [System.Text.StringBuilder]::new()
                $bProps = [ordered]@{
                    _DiffStatus     = $status
                    _ChangedColumns = $changed
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($col in $BaseAllProps) {
                    $v = $bRow.$col
                    $bProps[$col] = $v
                    $bProps[$baseSafeProps[$col]] = $changedBaseCols.Contains($col)
                    if ($null -ne $v) { [void]$sbBase.Append($v.ToString()); [void]$sbBase.Append('|') }
                }
                $bProps['_SearchableText'] = $sbBase.ToString()

                $sbUpd = [System.Text.StringBuilder]::new()
                $uProps = [ordered]@{
                    _DiffStatus     = $status
                    _ChangedColumns = $changed
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($col in $UpdateAllProps) {
                    $v = $uRow.$col
                    $uProps[$col] = $v
                    $uProps[$updateSafeProps[$col]] = $changedUpdateCols.Contains($col)
                    if ($null -ne $v) { [void]$sbUpd.Append($v.ToString()); [void]$sbUpd.Append('|') }
                }
                $uProps['_SearchableText'] = $sbUpd.ToString()

                [void]$resultBase.Add([PSCustomObject]$bProps)
                [void]$resultUpdate.Add([PSCustomObject]$uProps)
            }
        }

        return @{ Base = $resultBase; Update = $resultUpdate }
    }
    # EndRegion

    $baseIndex = @{}
    $updateIndex = @{}

    foreach ($row in $BaseData) {
        $key = Build-JoinKey -Row $row -KeyColumns $JoinBaseColumns -Trim $TrimWhitespace -IgnoreCase $IgnoreCase -IgnoreSpecialChars $IgnoreSpecialChars -IgnoreAllWhitespace $IgnoreAllSpaces
        if (-not $baseIndex.ContainsKey($key)) { $baseIndex[$key] = [System.Collections.Generic.List[object]]::new() }
        $baseIndex[$key].Add($row)
    }
    foreach ($row in $UpdateData) {
        $key = Build-JoinKey -Row $row -KeyColumns $JoinUpdateColumns -Trim $TrimWhitespace -IgnoreCase $IgnoreCase -IgnoreSpecialChars $IgnoreSpecialChars -IgnoreAllWhitespace $IgnoreAllSpaces
        if (-not $updateIndex.ContainsKey($key)) { $updateIndex[$key] = [System.Collections.Generic.List[object]]::new() }
        $updateIndex[$key].Add($row)
    }

    $allKeys = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($k in $baseIndex.Keys) { [void]$allKeys.Add($k) }
    foreach ($k in $updateIndex.Keys) { [void]$allKeys.Add($k) }

    $resultBase = [System.Collections.ArrayList]::new()
    $resultUpdate = [System.Collections.ArrayList]::new()

    $bucketDeleted   = [System.Collections.Generic.List[int]]::new()
    $bucketAdded     = [System.Collections.Generic.List[int]]::new()
    $bucketModified  = [System.Collections.Generic.List[int]]::new()
    $bucketUnchanged = [System.Collections.Generic.List[int]]::new()

    foreach ($key in $allKeys) {
        $baseRows = $null
        if ($baseIndex.ContainsKey($key)) { $baseRows = $baseIndex[$key] }
        $updateRows = $null
        if ($updateIndex.ContainsKey($key)) { $updateRows = $updateIndex[$key] }

        $bCount = 0
        if ($null -ne $baseRows) { $bCount = $baseRows.Count }
        $uCount = 0
        if ($null -ne $updateRows) { $uCount = $updateRows.Count }
        $maxPairs = [Math]::Max($bCount, $uCount)

        for ($i = 0; $i -lt $maxPairs; $i++) {
            $bRow = $null
            if ($null -ne $baseRows -and $i -lt $baseRows.Count) { $bRow = $baseRows[$i] }
            $uRow = $null
            if ($null -ne $updateRows -and $i -lt $updateRows.Count) { $uRow = $updateRows[$i] }

            $pairIdx = $resultBase.Count
            $pairId = "P_$pairIdx"

            if ($null -eq $bRow) {
                # Added - only in update
                $sbUpd = [System.Text.StringBuilder]::new()
                $uProps = [ordered]@{
                    _DiffStatus     = $statusAdded
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $UpdateAllProps) {
                    $v = $uRow.$p
                    $uProps[$p] = $v
                    $uProps[$updateSafeProps[$p]] = $false
                    if ($null -ne $v) { [void]$sbUpd.Append($v.ToString()); [void]$sbUpd.Append('|') }
                }
                $uProps['_SearchableText'] = $sbUpd.ToString()

                $blank = [ordered]@{
                    _DiffStatus     = $statusAdded
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $BaseAllProps) {
                    $blank[$p] = ''
                    $blank[$baseSafeProps[$p]] = $false
                }
                $blank['_SearchableText'] = ''

                [void]$resultBase.Add([PSCustomObject]$blank)
                [void]$resultUpdate.Add([PSCustomObject]$uProps)
                $bucketAdded.Add($pairIdx)
            }
            elseif ($null -eq $uRow) {
                # Deleted - only in base
                $sbBase = [System.Text.StringBuilder]::new()
                $bProps = [ordered]@{
                    _DiffStatus     = $statusDeleted
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $BaseAllProps) {
                    $v = $bRow.$p
                    $bProps[$p] = $v
                    $bProps[$baseSafeProps[$p]] = $false
                    if ($null -ne $v) { [void]$sbBase.Append($v.ToString()); [void]$sbBase.Append('|') }
                }
                $bProps['_SearchableText'] = $sbBase.ToString()

                $blank = [ordered]@{
                    _DiffStatus     = $statusDeleted
                    _ChangedColumns = ''
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($p in $UpdateAllProps) {
                    $blank[$p] = ''
                    $blank[$updateSafeProps[$p]] = $false
                }
                $blank['_SearchableText'] = ''

                [void]$resultBase.Add([PSCustomObject]$bProps)
                [void]$resultUpdate.Add([PSCustomObject]$blank)
                $bucketDeleted.Add($pairIdx)
            }
            else {
                # Both exist - check each mapping rule
                $changedLabels = [System.Collections.Generic.List[string]]::new()
                $changedBaseCols = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $changedUpdateCols = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

                foreach ($rule in $compiledRules) {
                    $isMatch = $false
                    # Fast-path for 1-to-1 direct mapping rules
                    if ($rule.IsDirect) {
                        $bValObj = $bRow.($rule.BaseCol)
                        $bRaw = if ($null -ne $bValObj) { $bValObj.ToString() } else { '' }
                        $uValObj = $uRow.($rule.UpdateCol)
                        $uRaw = if ($null -ne $uValObj) { $uValObj.ToString() } else { '' }
                        $isMatch = [FastDiffHelper]::AreEqual($bRaw, $uRaw, $IgnoreCase, $TrimWhitespace, $IgnoreSpecialChars, $IgnoreAllSpaces)
                    }
                    else {
                        $bVal = Get-MergedValue -Row $bRow -Columns $rule.BaseColumns   -MergeMode $rule.MergeMode -Separator $rule.Separator -Trim $TrimWhitespace -IgnoreCase $IgnoreCase -IgnoreSpecialChars $IgnoreSpecialChars -IgnoreAllWhitespace $IgnoreAllSpaces
                        $uVal = Get-MergedValue -Row $uRow -Columns $rule.UpdateColumns -MergeMode $rule.MergeMode -Separator $rule.Separator -Trim $TrimWhitespace -IgnoreCase $IgnoreCase -IgnoreSpecialChars $IgnoreSpecialChars -IgnoreAllWhitespace $IgnoreAllSpaces
                        $isMatch = ($bVal -eq $uVal)
                    }

                    if (-not $isMatch) {
                        $changedLabels.Add($rule.Label)
                        foreach ($c in $rule.BaseColumns) { [void]$changedBaseCols.Add($c) }
                        foreach ($c in $rule.UpdateColumns) { [void]$changedUpdateCols.Add($c) }
                    }
                }

                $status = if ($changedLabels.Count -gt 0) { $statusModified } else { $statusUnchanged }
                $changed = $changedLabels -join ', '

                $sbBase = [System.Text.StringBuilder]::new()
                $bProps = [ordered]@{
                    _DiffStatus     = $status
                    _ChangedColumns = $changed
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($col in $BaseAllProps) {
                    $v = $bRow.$col
                    $bProps[$col] = $v
                    $bProps[$baseSafeProps[$col]] = $changedBaseCols.Contains($col)
                    if ($null -ne $v) { [void]$sbBase.Append($v.ToString()); [void]$sbBase.Append('|') }
                }
                $bProps['_SearchableText'] = $sbBase.ToString()

                $sbUpd = [System.Text.StringBuilder]::new()
                $uProps = [ordered]@{
                    _DiffStatus     = $status
                    _ChangedColumns = $changed
                    _PairId         = $pairId
                    _JoinKey        = $key
                }
                foreach ($col in $UpdateAllProps) {
                    $v = $uRow.$col
                    $uProps[$col] = $v
                    $uProps[$updateSafeProps[$col]] = $changedUpdateCols.Contains($col)
                    if ($null -ne $v) { [void]$sbUpd.Append($v.ToString()); [void]$sbUpd.Append('|') }
                }
                $uProps['_SearchableText'] = $sbUpd.ToString()

                [void]$resultBase.Add([PSCustomObject]$bProps)
                [void]$resultUpdate.Add([PSCustomObject]$uProps)

                if ($status -eq $statusModified) {
                    $bucketModified.Add($pairIdx)
                }
                else {
                    $bucketUnchanged.Add($pairIdx)
                }
            }
        }
    }

    # Sort output: Deleted=0, Added=1, Modified=2, Unchanged=3 via O(N) bucket collection
    if ($resultBase.Count -eq 0) {
        return @{ Base = [System.Collections.ArrayList]::new(); Update = [System.Collections.ArrayList]::new() }
    }

    $sortedBase = [System.Collections.ArrayList]::new($resultBase.Count)
    $sortedUpdate = [System.Collections.ArrayList]::new($resultUpdate.Count)

    foreach ($idx in $bucketDeleted) {
        [void]$sortedBase.Add($resultBase[$idx])
        [void]$sortedUpdate.Add($resultUpdate[$idx])
    }
    foreach ($idx in $bucketAdded) {
        [void]$sortedBase.Add($resultBase[$idx])
        [void]$sortedUpdate.Add($resultUpdate[$idx])
    }
    foreach ($idx in $bucketModified) {
        [void]$sortedBase.Add($resultBase[$idx])
        [void]$sortedUpdate.Add($resultUpdate[$idx])
    }
    foreach ($idx in $bucketUnchanged) {
        [void]$sortedBase.Add($resultBase[$idx])
        [void]$sortedUpdate.Add($resultUpdate[$idx])
    }
    return @{ Base = $sortedBase; Update = $sortedUpdate }
}
#endregion


#region 4. Column Mapping Rule Dialog

function Show-MappingRuleDialog {
    param (
        [string[]]$BaseHeaders,
        [string[]]$UpdateHeaders,
        [hashtable]$ExistingRule = $null
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-Loc 'MapDialogTitle'
    $form.Size = New-Object System.Drawing.Size(620, 560)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $pal = Get-ThemePalette $Global:CurrentTheme
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($pal.BgCard)
    $form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)

    $lblBase = New-Object System.Windows.Forms.Label
    $lblBase.Text = Get-Loc 'MapBaseLabel'; $lblBase.Location = New-Object System.Drawing.Point(12, 14); $lblBase.AutoSize = $true

    $lbBase = New-Object System.Windows.Forms.ListBox
    $lbBase.SelectionMode = 'MultiExtended'
    $lbBase.Location = New-Object System.Drawing.Point(12, 36)
    $lbBase.Size = New-Object System.Drawing.Size(280, 155)
    $lbBase.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $lbBase.BackColor = [System.Drawing.ColorTranslator]::FromHtml($pal.BgInput)
    $lbBase.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)
    foreach ($h in $BaseHeaders) { [void]$lbBase.Items.Add($h) }

    $lblUpd = New-Object System.Windows.Forms.Label
    $lblUpd.Text = Get-Loc 'MapUpdateLabel'; $lblUpd.Location = New-Object System.Drawing.Point(310, 14); $lblUpd.AutoSize = $true

    $lbUpd = New-Object System.Windows.Forms.ListBox
    $lbUpd.SelectionMode = 'MultiExtended'
    $lbUpd.Location = New-Object System.Drawing.Point(310, 36)
    $lbUpd.Size = New-Object System.Drawing.Size(280, 155)
    $lbUpd.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $lbUpd.BackColor = [System.Drawing.ColorTranslator]::FromHtml($pal.BgInput)
    $lbUpd.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)
    foreach ($h in $UpdateHeaders) { [void]$lbUpd.Items.Add($h) }

    $lblMerge = New-Object System.Windows.Forms.Label
    $lblMerge.Text = Get-Loc 'MapMergeLabel'; $lblMerge.Location = New-Object System.Drawing.Point(12, 206); $lblMerge.AutoSize = $true

    $cmbMerge = New-Object System.Windows.Forms.ComboBox
    $cmbMerge.DropDownStyle = 'DropDownList'
    $cmbMerge.Location = New-Object System.Drawing.Point(12, 228); $cmbMerge.Width = 190
    $cmbMerge.BackColor = [System.Drawing.ColorTranslator]::FromHtml($pal.BgInput)
    $cmbMerge.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)
    @('Concatenate', 'Exact', 'FirstNonEmpty') | ForEach-Object { [void]$cmbMerge.Items.Add($_) }
    $cmbMerge.SelectedIndex = 0

    $lblSep = New-Object System.Windows.Forms.Label
    $lblSep.Text = Get-Loc 'MapSepLabel'; $lblSep.Location = New-Object System.Drawing.Point(220, 206); $lblSep.AutoSize = $true

    $txtSep = New-Object System.Windows.Forms.TextBox
    $txtSep.Location = New-Object System.Drawing.Point(220, 228); $txtSep.Width = 160; $txtSep.Text = ' | '
    $txtSep.BackColor = [System.Drawing.ColorTranslator]::FromHtml($pal.BgInput)
    $txtSep.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)

    $lblLabel = New-Object System.Windows.Forms.Label
    $lblLabel.Text = Get-Loc 'MapLabelLabel'; $lblLabel.Location = New-Object System.Drawing.Point(12, 268); $lblLabel.AutoSize = $true

    $txtLabel = New-Object System.Windows.Forms.TextBox
    $txtLabel.Location = New-Object System.Drawing.Point(12, 290); $txtLabel.Width = 578
    $txtLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml($pal.BgInput)
    $txtLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)

    # Preview area
    $grpPreview = New-Object System.Windows.Forms.GroupBox
    $grpPreview.Text = 'Preview'; $grpPreview.Location = New-Object System.Drawing.Point(12, 325)
    $grpPreview.Size = New-Object System.Drawing.Size(578, 90)
    $grpPreview.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextSecondary)

    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Location = New-Object System.Drawing.Point(8, 18)
    $lblPreview.Size = New-Object System.Drawing.Size(560, 64)
    $lblPreview.Text = 'Select columns above to see a preview...'
    $lblPreview.Font = New-Object System.Drawing.Font('Consolas', 9)
    $lblPreview.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)
    $grpPreview.Controls.Add($lblPreview)

    $UpdatePreview = {
        $bCols = @($lbBase.SelectedItems)
        $uCols = @($lbUpd.SelectedItems)
        $merge = $cmbMerge.SelectedItem
        $sep = $txtSep.Text
        $bStr = if ($bCols.Count -gt 0) { $bCols -join " $($sep.Trim()) " } else { '(none)' }
        $uStr = if ($uCols.Count -gt 0) { $uCols -join " $($sep.Trim()) " } else { '(none)' }
        $lblPreview.Text = "Base:   $bStr`r`nUpdate: $uStr`r`nMode:   $merge"
        if ($txtLabel.Text -eq '' -and $bCols.Count -gt 0 -and $uCols.Count -gt 0) {
            $txtLabel.Text = ($bCols -join ' | ') + ' <-> ' + ($uCols -join ' | ')
        }
    }

    $lbBase.add_SelectedIndexChanged($UpdatePreview)
    $lbUpd.add_SelectedIndexChanged($UpdatePreview)
    $cmbMerge.add_SelectedIndexChanged($UpdatePreview)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = Get-Loc 'MapOK'
    $btnOK.Location = New-Object System.Drawing.Point(400, 430)
    $btnOK.Size = New-Object System.Drawing.Size(90, 36)
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(16, 185, 129)
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.FlatStyle = 'Flat'
    $btnOK.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = Get-Loc 'MapCancel'
    $btnCancel.Location = New-Object System.Drawing.Point(502, 430)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 36)
    $btnCancel.DialogResult = 'Cancel'
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml($pal.BgButtonDefault)
    $btnCancel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($pal.TextPrimary)

    if ($null -ne $ExistingRule) {
        foreach ($col in $ExistingRule.BaseColumns) {
            $idx = $lbBase.Items.IndexOf($col); if ($idx -ge 0) { $lbBase.SetSelected($idx, $true) }
        }
        foreach ($col in $ExistingRule.UpdateColumns) {
            $idx = $lbUpd.Items.IndexOf($col); if ($idx -ge 0) { $lbUpd.SetSelected($idx, $true) }
        }
        $mi = @('Concatenate', 'Exact', 'FirstNonEmpty').IndexOf($ExistingRule.MergeMode)
        if ($mi -ge 0) { $cmbMerge.SelectedIndex = $mi }
        $txtSep.Text = $ExistingRule.Separator
        $txtLabel.Text = $ExistingRule.Label
    }

    $script:MapDlgResult = $null
    $btnOK.add_Click({
            $bCols = @($lbBase.SelectedItems)
            $uCols = @($lbUpd.SelectedItems)
            if ($bCols.Count -eq 0 -or $uCols.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show('Select at least one base AND one update column.', 'Warning', 0, 48)
                return
            }
            $lbl = $txtLabel.Text.Trim()
            if ($lbl -eq '') { $lbl = ($bCols -join ' | ') + ' <-> ' + ($uCols -join ' | ') }
            $script:MapDlgResult = @{
                BaseColumns   = $bCols
                UpdateColumns = $uCols
                MergeMode     = $cmbMerge.SelectedItem.ToString()
                Separator     = $txtSep.Text
                Label         = $lbl
            }
            $form.DialogResult = 'OK'; $form.Close()
        })

    $form.Controls.AddRange(@($lblBase, $lbBase, $lblUpd, $lbUpd, $lblMerge, $cmbMerge, $lblSep, $txtSep, $lblLabel, $txtLabel, $grpPreview, $btnOK, $btnCancel))
    $form.AcceptButton = $btnOK; $form.CancelButton = $btnCancel
    $form.add_Shown({
            Set-WindowDarkMode -WindowHandle $form.Handle -IsDark $pal.IsDark
        })
    $form.ShowDialog() | Out-Null
    return $script:MapDlgResult
}
#endregion

#region 5. Results Viewer

function Show-CompareResult {
    param (
        [System.Collections.ArrayList]$DiffBase,
        [System.Collections.ArrayList]$DiffUpdate,
        [string[]]$BaseAllProps,
        [string[]]$UpdateAllProps,
        [string]$BaseFileName,
        [string]$UpdateFileName,
        [hashtable[]]$MappingRules,
        [string]$DefaultLayout = 'Vertical'
    )

    $statusAdded = Get-Loc 'StatusAdded'
    $statusDeleted = Get-Loc 'StatusDeleted'
    $statusModified = Get-Loc 'StatusModified'
    $statusUnchanged = Get-Loc 'StatusUnchanged'
    $statusAll = Get-Loc 'StatusAll'

    # Ensure _PairId and _SearchableText exist if not already populated by fast diff engine
    if ($DiffBase.Count -gt 0) {
        $hasPairId = ($DiffBase[0].PSObject.Properties.Match('_PairId').Count -gt 0)
        $hasSearch = ($DiffBase[0].PSObject.Properties.Match('_SearchableText').Count -gt 0)
        if (-not $hasPairId -or -not $hasSearch) {
            for ($i = 0; $i -lt $DiffBase.Count; $i++) {
                $b = $DiffBase[$i]
                $u = $DiffUpdate[$i]
                $pairId = if ($hasPairId -and $null -ne $b._PairId) { $b._PairId } else { "P_$i" }
                if (-not $hasPairId) {
                    $b | Add-Member -MemberType NoteProperty -Name '_PairId' -Value $pairId -Force
                    $u | Add-Member -MemberType NoteProperty -Name '_PairId' -Value $pairId -Force
                }
                if (-not $hasSearch) {
                    $sbB = [System.Text.StringBuilder]::new()
                    foreach ($p in $b.PSObject.Properties.Name) {
                        if (-not $p.StartsWith('_') -and $null -ne $b.$p) { [void]$sbB.Append($b.$p.ToString()); [void]$sbB.Append('|') }
                    }
                    $b | Add-Member -MemberType NoteProperty -Name '_SearchableText' -Value $sbB.ToString() -Force

                    $sbU = [System.Text.StringBuilder]::new()
                    foreach ($p in $u.PSObject.Properties.Name) {
                        if (-not $p.StartsWith('_') -and $null -ne $u.$p) { [void]$sbU.Append($u.$p.ToString()); [void]$sbU.Append('|') }
                    }
                    $u | Add-Member -MemberType NoteProperty -Name '_SearchableText' -Value $sbU.ToString() -Force
                }
            }
        }
    }

    $script:ViewBase = $DiffBase
    $script:ViewUpdate = $DiffUpdate
    $Title = "$BaseFileName  <->  $UpdateFileName"

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Diff Results"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource BgWindowBrush}"
        WindowState="Maximized"
        Topmost="True">
    <Window.Resources>
        <SolidColorBrush x:Key="BgWindowBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgHeaderBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgCardBrush" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgInputBrush" Color="#334155"/>
        <SolidColorBrush x:Key="BgButtonDefaultBrush" Color="#334155"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#475569"/>
        <SolidColorBrush x:Key="BorderSubtleBrush" Color="#334155"/>
        <SolidColorBrush x:Key="GridLinesBrush" Color="#334155"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#CBD5E1"/>
        <SolidColorBrush x:Key="TextMutedBrush" Color="#94A3B8"/>
        <SolidColorBrush x:Key="RowBgBrush" Color="#1E293B"/>
        <SolidColorBrush x:Key="RowAltBgBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="HeaderBgBrush" Color="#334155"/>
        <SolidColorBrush x:Key="HeaderFgBrush" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="SummaryBgBrush" Color="#020617"/>
        <SolidColorBrush x:Key="ItemHoverBgBrush" Color="#334155"/>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource BgInputBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton"
                                          Focusable="false"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border x:Name="Border" CornerRadius="4" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                                            <Path x:Name="Arrow" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" Fill="{DynamicResource TextSecondaryBrush}" Data="M 0 0 L 4 4 L 8 0 Z"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite"
                                              IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="8,4,24,4"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"
                                              TextElement.Foreground="{TemplateBinding Foreground}"/>
                            <Popup x:Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True"
                                   Focusable="False"
                                   PopupAnimation="Slide">
                                <Border x:Name="DropDownBorder"
                                        Background="{DynamicResource BgCardBrush}"
                                        BorderBrush="{DynamicResource BorderBrush}"
                                        BorderThickness="1"
                                        CornerRadius="4"
                                        MinWidth="{TemplateBinding ActualWidth}"
                                        MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <ScrollViewer Margin="2" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Border" Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource ItemHoverBgBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource BgCardBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="RowBackground" Value="{DynamicResource RowBgBrush}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource RowAltBgBrush}"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource GridLinesBrush}"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Background" Value="Transparent"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#2563EB"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
        </Style>
    </Window.Resources>
    <Grid x:Name="mainGrid" Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0"
                   FontSize="19" FontWeight="SemiBold" Foreground="{DynamicResource TextPrimaryBrush}" Margin="0,0,0,10"
                   x:Name="lblTitle"/>

        <Border x:Name="borderFilters" Grid.Row="1" Background="{DynamicResource BgCardBrush}" CornerRadius="6"
                BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1"
                Padding="12,8" Margin="0,0,0,10">
            <WrapPanel Orientation="Horizontal" VerticalAlignment="Center">
                <StackPanel Margin="0,0,12,4">
                    <TextBlock x:Name="lblLayout" Text="Layout:" Foreground="{DynamicResource TextSecondaryBrush}" Margin="0,0,0,2" FontSize="11"/>
                    <ComboBox x:Name="cmbLayout" Width="175" Padding="3"/>
                </StackPanel>
                <StackPanel Margin="0,0,12,4">
                    <TextBlock x:Name="lblStatus" Text="Status:" Foreground="{DynamicResource TextSecondaryBrush}" Margin="0,0,0,2" FontSize="11"/>
                    <ComboBox x:Name="cmbStatus" Width="130" Padding="3"/>
                </StackPanel>
                <StackPanel Margin="0,0,12,4">
                    <TextBlock x:Name="lblSearch" Text="Search:" Foreground="{DynamicResource TextSecondaryBrush}" Margin="0,0,0,2" FontSize="11"/>
                    <TextBox x:Name="txtSearch" Width="200" Padding="6,4" Background="{DynamicResource BgInputBrush}" Foreground="{DynamicResource TextPrimaryBrush}" BorderBrush="{DynamicResource BorderBrush}"/>
                </StackPanel>
                <Button x:Name="btnReset" Content="Reset" Width="90" Height="28" Margin="0,15,8,0"
                        Background="{DynamicResource BgButtonDefaultBrush}" Foreground="{DynamicResource TextPrimaryBrush}" BorderThickness="0" Cursor="Hand"/>
                <Button x:Name="btnClearRowFilter" Content="✕ Clear Row Filter" Height="28" Padding="10,0" Margin="0,15,8,0"
                        Background="#DC2626" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand" Visibility="Collapsed"/>
                <Button x:Name="btnResTheme" Content="Theme" Width="90" Height="28" Margin="0,15,8,0"
                        Background="{DynamicResource BgButtonDefaultBrush}" Foreground="{DynamicResource TextPrimaryBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" Cursor="Hand" FontWeight="SemiBold" FontSize="12"/>
                <CheckBox x:Name="chkSync" Content="Sync" Foreground="{DynamicResource TextPrimaryBrush}" Margin="4,18,12,0"
                          FontSize="12" IsChecked="True" VerticalAlignment="Center"/>
                <Button x:Name="btnExport" Content="Export" Width="120" Height="28" Margin="0,15,0,0"
                        Background="#2563EB" Foreground="White" FontWeight="Bold"
                        BorderThickness="0" Cursor="Hand"/>
            </WrapPanel>
        </Border>

        <Border x:Name="borderBase" Grid.Row="2" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1"
                Margin="0,0,0,5" CornerRadius="4">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Border x:Name="borderBaseTitle" Background="#1F2937" CornerRadius="4,4,0,0" Padding="10,6">
                    <TextBlock x:Name="lblBase" Foreground="White" FontWeight="SemiBold" FontSize="12"/>
                </Border>
                <DataGrid x:Name="dgBase" Grid.Row="1"
                          AutoGenerateColumns="True" IsReadOnly="True" CanUserSortColumns="True"
                          SelectionMode="Single" SelectionUnit="FullRow"
                          RowHeaderWidth="0" GridLinesVisibility="Horizontal"
                          HeadersVisibility="Column" BorderThickness="0">
                    <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                            <Setter Property="FontWeight" Value="Bold"/>
                            <Setter Property="Background" Value="{DynamicResource HeaderBgBrush}"/>
                            <Setter Property="Foreground" Value="{DynamicResource HeaderFgBrush}"/>
                            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
                            <Setter Property="BorderThickness" Value="0,0,1,1"/>
                            <Setter Property="Padding" Value="6,4"/>
                            <Setter Property="Cursor" Value="Hand"/>
                        </Style>
                    </DataGrid.ColumnHeaderStyle>
                </DataGrid>
            </Grid>
        </Border>

        <Border x:Name="borderUpdate" Grid.Row="3" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1"
                Margin="0,5,0,0" CornerRadius="4">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Border x:Name="borderUpdateTitle" Background="#0D9488" CornerRadius="4,4,0,0" Padding="10,6">
                    <TextBlock x:Name="lblUpdate" Foreground="White" FontWeight="SemiBold" FontSize="12"/>
                </Border>
                <DataGrid x:Name="dgUpdate" Grid.Row="1"
                          AutoGenerateColumns="True" IsReadOnly="True" CanUserSortColumns="True"
                          SelectionMode="Single" SelectionUnit="FullRow"
                          RowHeaderWidth="0" GridLinesVisibility="Horizontal"
                          HeadersVisibility="Column" BorderThickness="0">
                    <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                            <Setter Property="FontWeight" Value="Bold"/>
                            <Setter Property="Background" Value="{DynamicResource HeaderBgBrush}"/>
                            <Setter Property="Foreground" Value="{DynamicResource HeaderFgBrush}"/>
                            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
                            <Setter Property="BorderThickness" Value="0,0,1,1"/>
                            <Setter Property="Padding" Value="6,4"/>
                            <Setter Property="Cursor" Value="Hand"/>
                        </Style>
                    </DataGrid.ColumnHeaderStyle>
                </DataGrid>
            </Grid>
        </Border>

        <Border x:Name="borderSummary" Grid.Row="4" Background="{DynamicResource SummaryBgBrush}" CornerRadius="6"
                Padding="14,8" Margin="0,10,0,0">
            <WrapPanel x:Name="pnlSummary" Orientation="Horizontal" HorizontalAlignment="Center"/>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $Window = [Windows.Markup.XamlReader]::Load($reader)

    $mainGrid = $Window.FindName('mainGrid')
    $borderFilters = $Window.FindName('borderFilters')
    $borderBase = $Window.FindName('borderBase')
    $borderBaseTitle = $Window.FindName('borderBaseTitle')
    $borderUpdate = $Window.FindName('borderUpdate')
    $borderUpdateTitle = $Window.FindName('borderUpdateTitle')
    $borderSummary = $Window.FindName('borderSummary')
    $lblLayout = $Window.FindName('lblLayout')
    $lblStatus = $Window.FindName('lblStatus')
    $lblSearch = $Window.FindName('lblSearch')
    $cmbLayout = $Window.FindName('cmbLayout')
    $dgBase = $Window.FindName('dgBase')
    $dgUpdate = $Window.FindName('dgUpdate')
    $cmbStatus = $Window.FindName('cmbStatus')
    $txtSearch = $Window.FindName('txtSearch')
    $btnReset = $Window.FindName('btnReset')
    $btnClearRowFilter = $Window.FindName('btnClearRowFilter')
    $btnResTheme = $Window.FindName('btnResTheme')
    $chkSync = $Window.FindName('chkSync')
    $btnExport = $Window.FindName('btnExport')
    $pnlSummary = $Window.FindName('pnlSummary')
    $lblTitle = $Window.FindName('lblTitle')
    $lblBase = $Window.FindName('lblBase')
    $lblUpdate = $Window.FindName('lblUpdate')

    $btnReset.Content = Get-Loc 'DashReset'
    if ($btnClearRowFilter) { $btnClearRowFilter.Content = Get-Loc 'BtnClearRowFilter' }
    $chkSync.Content = Get-Loc 'SyncScroll'
    $btnExport.Content = Get-Loc 'DashExport'
    $lblTitle.Text = $Title
    $lblBase.Text = "$(Get-Loc 'PanelBase'): $BaseFileName"
    $lblUpdate.Text = "$(Get-Loc 'PanelUpdate'): $UpdateFileName"

    if ($lblLayout) { $lblLayout.Text = Get-Loc 'LayoutMode' }
    if ($lblStatus) { $lblStatus.Text = "$(Get-Loc 'ColDiffStatus'):" }
    if ($lblSearch) { $lblSearch.Text = Get-Loc 'DashSearch' }

    $ApplyViewerTheme = {
        $pal = Get-ThemePalette $Global:CurrentTheme
        Update-WpfThemeResources $Window $pal

        if ($btnResTheme) {
            $btnResTheme.Content = if ($pal.IsDark) { Get-Loc 'ThemeLight' } else { Get-Loc 'ThemeDark' }
        }

        $tipText = [System.Security.SecurityElement]::Escape((Get-Loc 'TipRowFilter'))
        $rXaml = @"
<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
       TargetType="DataGridRow">
    <Setter Property="Foreground" Value="$($pal.TextPrimary)"/>
    <Setter Property="ToolTip" Value="$tipText"/>
    <Style.Triggers>
        <DataTrigger Binding="{Binding _DiffStatus}" Value="$statusAdded">
            <Setter Property="Background" Value="$($pal.RowAddedBg)"/>
            <Setter Property="Foreground" Value="$($pal.RowAddedFg)"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding _DiffStatus}" Value="$statusDeleted">
            <Setter Property="Background" Value="$($pal.RowDeletedBg)"/>
            <Setter Property="Foreground" Value="$($pal.RowDeletedFg)"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding _DiffStatus}" Value="$statusModified">
            <Setter Property="Background" Value="$($pal.RowModifiedBg)"/>
            <Setter Property="Foreground" Value="$($pal.RowModifiedFg)"/>
        </DataTrigger>
    </Style.Triggers>
</Style>
"@
        $rStyle = [System.Windows.Markup.XamlReader]::Parse($rXaml)
        $dgBase.RowStyle = $rStyle
        $dgUpdate.RowStyle = $rStyle

        # Re-apply cell highlight styles for columns
        foreach ($col in $dgBase.Columns) {
            $h = $col.Header.ToString()
            if (-not $h.StartsWith('_')) {
                $pName = Get-SafePropName $h
                $cXaml = @"
<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
       TargetType="DataGridCell">
    <Setter Property="Foreground" Value="$($pal.TextPrimary)"/>
    <Style.Triggers>
        <DataTrigger Binding="{Binding $pName}" Value="True">
            <Setter Property="Background" Value="$($pal.CellChgBg)"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderBrush" Value="$($pal.CellChgBorder)"/>
            <Setter Property="BorderThickness" Value="1.5"/>
            <Setter Property="Foreground" Value="$($pal.CellChgFg)"/>
        </DataTrigger>
    </Style.Triggers>
</Style>
"@
                try { $col.CellStyle = [System.Windows.Markup.XamlReader]::Parse($cXaml) } catch {}
            }
        }
        foreach ($col in $dgUpdate.Columns) {
            $h = $col.Header.ToString()
            if (-not $h.StartsWith('_')) {
                $pName = Get-SafePropName $h
                $cXaml = @"
<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
       TargetType="DataGridCell">
    <Setter Property="Foreground" Value="$($pal.TextPrimary)"/>
    <Style.Triggers>
        <DataTrigger Binding="{Binding $pName}" Value="True">
            <Setter Property="Background" Value="$($pal.CellChgBg)"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderBrush" Value="$($pal.CellChgBorder)"/>
            <Setter Property="BorderThickness" Value="1.5"/>
            <Setter Property="Foreground" Value="$($pal.CellChgFg)"/>
        </DataTrigger>
    </Style.Triggers>
</Style>
"@
                try { $col.CellStyle = [System.Windows.Markup.XamlReader]::Parse($cXaml) } catch {}
            }
        }

        Set-WindowDarkMode -Window $Window -IsDark $pal.IsDark
    }

    $btnResTheme.add_Click({
            $Global:CurrentTheme = if ($Global:CurrentTheme -eq 'Dark') { 'Light' } else { 'Dark' }
            try { Save-AppSettings } catch {}
            & $ApplyViewerTheme
        })

    $Window.add_Loaded({
            & $ApplyViewerTheme
        })

    # Layout options
    $itemVert = New-Object System.Windows.Controls.ComboBoxItem; $itemVert.Content = Get-Loc 'LayoutVert'; $itemVert.Tag = 'Vertical'
    $itemHoriz = New-Object System.Windows.Controls.ComboBoxItem; $itemHoriz.Content = Get-Loc 'LayoutHoriz'; $itemHoriz.Tag = 'Horizontal'
    [void]$cmbLayout.Items.Add($itemVert)
    [void]$cmbLayout.Items.Add($itemHoriz)
    $cmbLayout.SelectedIndex = if ($DefaultLayout -eq 'Horizontal') { 1 } else { 0 }

    $SetGridLayout = {
        param([string]$Mode)
        $mainGrid.RowDefinitions.Clear()
        $mainGrid.ColumnDefinitions.Clear()

        if ($Mode -eq 'Horizontal') {
            # Side by Side
            @(
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }),
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }),
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = (New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)) }),
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto })
            ) | ForEach-Object { $mainGrid.RowDefinitions.Add($_) }

            @(
                (New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = (New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)) }),
                (New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = (New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)) })
            ) | ForEach-Object { $mainGrid.ColumnDefinitions.Add($_) }

            [System.Windows.Controls.Grid]::SetRow($lblTitle, 0); [System.Windows.Controls.Grid]::SetColumn($lblTitle, 0); [System.Windows.Controls.Grid]::SetColumnSpan($lblTitle, 2)
            [System.Windows.Controls.Grid]::SetRow($borderFilters, 1); [System.Windows.Controls.Grid]::SetColumn($borderFilters, 0); [System.Windows.Controls.Grid]::SetColumnSpan($borderFilters, 2)
            [System.Windows.Controls.Grid]::SetRow($borderBase, 2); [System.Windows.Controls.Grid]::SetColumn($borderBase, 0); [System.Windows.Controls.Grid]::SetColumnSpan($borderBase, 1)
            [System.Windows.Controls.Grid]::SetRow($borderUpdate, 2); [System.Windows.Controls.Grid]::SetColumn($borderUpdate, 1); [System.Windows.Controls.Grid]::SetColumnSpan($borderUpdate, 1)
            [System.Windows.Controls.Grid]::SetRow($borderSummary, 3); [System.Windows.Controls.Grid]::SetColumn($borderSummary, 0); [System.Windows.Controls.Grid]::SetColumnSpan($borderSummary, 2)

            $borderBase.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
            $borderUpdate.Margin = New-Object System.Windows.Thickness(5, 0, 0, 0)
        }
        else {
            # Top / Bottom (Vertical)
            @(
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }),
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto }),
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = (New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)) }),
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = (New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)) }),
                (New-Object System.Windows.Controls.RowDefinition -Property @{ Height = [System.Windows.GridLength]::Auto })
            ) | ForEach-Object { $mainGrid.RowDefinitions.Add($_) }

            $mainGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = (New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)) }))

            [System.Windows.Controls.Grid]::SetRow($lblTitle, 0); [System.Windows.Controls.Grid]::SetColumn($lblTitle, 0); [System.Windows.Controls.Grid]::SetColumnSpan($lblTitle, 1)
            [System.Windows.Controls.Grid]::SetRow($borderFilters, 1); [System.Windows.Controls.Grid]::SetColumn($borderFilters, 0); [System.Windows.Controls.Grid]::SetColumnSpan($borderFilters, 1)
            [System.Windows.Controls.Grid]::SetRow($borderBase, 2); [System.Windows.Controls.Grid]::SetColumn($borderBase, 0); [System.Windows.Controls.Grid]::SetColumnSpan($borderBase, 1)
            [System.Windows.Controls.Grid]::SetRow($borderUpdate, 3); [System.Windows.Controls.Grid]::SetColumn($borderUpdate, 0); [System.Windows.Controls.Grid]::SetColumnSpan($borderUpdate, 1)
            [System.Windows.Controls.Grid]::SetRow($borderSummary, 4); [System.Windows.Controls.Grid]::SetColumn($borderSummary, 0); [System.Windows.Controls.Grid]::SetColumnSpan($borderSummary, 1)

            $borderBase.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
            $borderUpdate.Margin = New-Object System.Windows.Thickness(0, 5, 0, 0)
        }
    }

    $cmbLayout.add_SelectionChanged({
            if ($cmbLayout.SelectedItem -and $cmbLayout.SelectedItem.Tag) {
                & $SetGridLayout $cmbLayout.SelectedItem.Tag
            }
        })

    $initialLayout = if ($DefaultLayout -eq 'Horizontal') { 'Horizontal' } else { 'Vertical' }
    & $SetGridLayout $initialLayout

    # Populate status combobox
    @($statusAll, $statusAdded, $statusDeleted, $statusModified, $statusUnchanged) | ForEach-Object {
        $item = New-Object System.Windows.Controls.ComboBoxItem; $item.Content = $_
        [void]$cmbStatus.Items.Add($item)
    }
    $cmbStatus.SelectedIndex = 0

    $AutoCol = {
        param($sender, $e)
        $colHeader = $e.Column.Header.ToString()
        if ($colHeader.StartsWith('_')) {
            $e.Cancel = $true
            return
        }
        $e.Column.SortMemberPath = $colHeader
        $e.Column.CanUserSort = $true

        $propName = Get-SafePropName $colHeader
        $pal = Get-ThemePalette $Global:CurrentTheme
        $cellStyleXaml = @"
<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
       TargetType="DataGridCell">
    <Style.Triggers>
        <DataTrigger Binding="{Binding $propName}" Value="True">
            <Setter Property="Background" Value="$($pal.CellChgBg)"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderBrush" Value="$($pal.CellChgBorder)"/>
            <Setter Property="BorderThickness" Value="1.5"/>
            <Setter Property="Foreground" Value="$($pal.CellChgFg)"/>
        </DataTrigger>
    </Style.Triggers>
</Style>
"@
        try {
            $e.Column.CellStyle = [System.Windows.Markup.XamlReader]::Parse($cellStyleXaml)
        }
        catch {}
    }
    $dgBase.add_AutoGeneratingColumn($AutoCol)
    $dgUpdate.add_AutoGeneratingColumn($AutoCol)

    # Summary bar
    $script:SummaryBlocks = @{}
    function Add-SumItem ([string]$K, [string]$Lbl, [string]$Clr = 'White') {
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = 'Horizontal'; $sp.Margin = '0,0,24,0'
        $t = New-Object System.Windows.Controls.TextBlock; $t.Text = "${Lbl}: "; $t.Foreground = '#9CA3AF'; $t.FontSize = 13
        $v = New-Object System.Windows.Controls.TextBlock; $v.Text = '0'; $v.Foreground = $Clr; $v.FontSize = 13
        $v.FontWeight = [System.Windows.FontWeights]::Bold
        $sp.Children.Add($t) | Out-Null; $sp.Children.Add($v) | Out-Null
        $pnlSummary.Children.Add($sp) | Out-Null
        $script:SummaryBlocks[$K] = $v
    }
    Add-SumItem 'Total'     (Get-Loc 'SummaryTotal')     '#F9FAFB'
    Add-SumItem 'Added'     (Get-Loc 'SummaryAdded')     '#86EFAC'
    Add-SumItem 'Deleted'   (Get-Loc 'SummaryDeleted')   '#FCA5A5'
    Add-SumItem 'Modified'  (Get-Loc 'SummaryModified')  '#FDE68A'
    Add-SumItem 'Unchanged' (Get-Loc 'SummaryUnchanged') '#9CA3AF'

    $script:ActiveFilterPairId = $null
    $script:PreviousSelectedPairId = $null
    $script:MouseDownRowBase = $null
    $script:MouseDownRowUpdate = $null
    $script:IgnoreSelectionChange = $false

    # Filter view
    $FilterView = {
        $statusFilter = $cmbStatus.SelectedItem.Content.ToString()
        $search = $txtSearch.Text.Trim()
        $hasSearch = $search.Length -ge 3

        if ($script:ViewBase.Count -eq 0) {
            $script:IgnoreSelectionChange = $true
            try {
                $dgBase.ItemsSource = [System.Collections.ArrayList]::new()
                $dgUpdate.ItemsSource = [System.Collections.ArrayList]::new()
            }
            finally {
                $script:IgnoreSelectionChange = $false
            }
            if ($script:SummaryBlocks.ContainsKey('Total')) {
                $script:SummaryBlocks['Total'].Text = '0'
                $script:SummaryBlocks['Added'].Text = '0'
                $script:SummaryBlocks['Deleted'].Text = '0'
                $script:SummaryBlocks['Modified'].Text = '0'
                $script:SummaryBlocks['Unchanged'].Text = '0'
            }
            if ($btnClearRowFilter) { $btnClearRowFilter.Visibility = [System.Windows.Visibility]::Collapsed }
            return
        }

        $indices = @(for ($i = 0; $i -lt $script:ViewBase.Count; $i++) { $i })
        if ($statusFilter -ne $statusAll) {
            $indices = @($indices | Where-Object { $script:ViewBase[$_]._DiffStatus -eq $statusFilter })
        }
        if ($hasSearch) {
            $indices = @($indices | Where-Object {
                    $script:ViewBase[$_]._SearchableText.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                    $script:ViewUpdate[$_]._SearchableText.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                })
        }
        if ($null -ne $script:ActiveFilterPairId) {
            $indices = @($indices | Where-Object {
                $script:ViewBase[$_]._PairId -eq $script:ActiveFilterPairId
            })
        }

        $filtB = @($indices | ForEach-Object { $script:ViewBase[$_] })
        $filtU = @($indices | ForEach-Object { $script:ViewUpdate[$_] })

        $script:IgnoreSelectionChange = $true
        try {
            $dgBase.ItemsSource = [System.Collections.ArrayList]@($filtB)
            $dgUpdate.ItemsSource = [System.Collections.ArrayList]@($filtU)
        }
        finally {
            $script:IgnoreSelectionChange = $false
        }

        if ($null -ne $script:ActiveFilterPairId -and $filtB.Count -gt 0) {
            $script:IgnoreSelectionChange = $true
            try {
                $dgBase.SelectedIndex = 0
                $dgUpdate.SelectedIndex = 0
            }
            finally {
                $script:IgnoreSelectionChange = $false
            }
            if ($btnClearRowFilter) { $btnClearRowFilter.Visibility = [System.Windows.Visibility]::Visible }
        }
        else {
            if ($btnClearRowFilter) { $btnClearRowFilter.Visibility = [System.Windows.Visibility]::Collapsed }
            if ($null -ne $script:PreviousSelectedPairId) {
                $prevId = $script:PreviousSelectedPairId
                $restoreIdx = -1
                for ($k = 0; $k -lt $filtB.Count; $k++) {
                    if ($filtB[$k]._PairId -eq $prevId) { $restoreIdx = $k; break }
                }
                if ($restoreIdx -ge 0) {
                    $script:IgnoreSelectionChange = $true
                    try {
                        $dgBase.SelectedIndex = $restoreIdx
                        $dgUpdate.SelectedIndex = $restoreIdx
                        $dgBase.ScrollIntoView($dgBase.SelectedItem)
                        $dgUpdate.ScrollIntoView($dgUpdate.SelectedItem)
                    }
                    finally {
                        $script:IgnoreSelectionChange = $false
                    }
                }
                $script:PreviousSelectedPairId = $null
            }
        }

        $script:SummaryBlocks['Total'].Text = $filtB.Count.ToString('N0')
        $script:SummaryBlocks['Added'].Text = (@($filtB | Where-Object { $_._DiffStatus -eq $statusAdded })).Count.ToString('N0')
        $script:SummaryBlocks['Deleted'].Text = (@($filtB | Where-Object { $_._DiffStatus -eq $statusDeleted })).Count.ToString('N0')
        $script:SummaryBlocks['Modified'].Text = (@($filtB | Where-Object { $_._DiffStatus -eq $statusModified })).Count.ToString('N0')
        $script:SummaryBlocks['Unchanged'].Text = (@($filtB | Where-Object { $_._DiffStatus -eq $statusUnchanged })).Count.ToString('N0')
    }

    $ToggleRowFilter = {
        param($item)
        if ($null -eq $item -or -not ($item.PSObject.Properties.Match('_PairId').Count -gt 0)) { return }
        $pairId = $item._PairId
        if ($script:ActiveFilterPairId -eq $pairId) {
            # Already filtered to this pair -> unfilter!
            $script:PreviousSelectedPairId = $pairId
            $script:ActiveFilterPairId = $null
        }
        else {
            # Filter to this pair!
            $script:PreviousSelectedPairId = $pairId
            $script:ActiveFilterPairId = $pairId
        }
        & $FilterView
    }

    $dgBase.add_PreviewMouseLeftButtonDown({
        param($s, $e)
        $script:MouseDownRowBase = Get-VisualParentRow $e.OriginalSource
    })
    $dgBase.add_PreviewMouseLeftButtonUp({
        param($s, $e)
        $mouseUpRow = Get-VisualParentRow $e.OriginalSource
        if ($null -ne $mouseUpRow -and $null -ne $script:MouseDownRowBase -and $mouseUpRow -eq $script:MouseDownRowBase) {
            & $ToggleRowFilter $mouseUpRow.Item
        }
        $script:MouseDownRowBase = $null
    })

    $dgUpdate.add_PreviewMouseLeftButtonDown({
        param($s, $e)
        $script:MouseDownRowUpdate = Get-VisualParentRow $e.OriginalSource
    })
    $dgUpdate.add_PreviewMouseLeftButtonUp({
        param($s, $e)
        $mouseUpRow = Get-VisualParentRow $e.OriginalSource
        if ($null -ne $mouseUpRow -and $null -ne $script:MouseDownRowUpdate -and $mouseUpRow -eq $script:MouseDownRowUpdate) {
            & $ToggleRowFilter $mouseUpRow.Item
        }
        $script:MouseDownRowUpdate = $null
    })

    $RowKeyDown = {
        param($grid, $e)
        if ($e.Key -in @([System.Windows.Input.Key]::Enter, [System.Windows.Input.Key]::Space)) {
            if ($grid.SelectedItem) {
                $e.Handled = $true
                & $ToggleRowFilter $grid.SelectedItem
            }
        }
    }
    $dgBase.add_KeyDown({ param($s, $e) & $RowKeyDown $dgBase $e })
    $dgUpdate.add_KeyDown({ param($s, $e) & $RowKeyDown $dgUpdate $e })

    $Window.add_KeyDown({
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
            if ($null -ne $script:ActiveFilterPairId) {
                $e.Handled = $true
                $script:ActiveFilterPairId = $null
                & $FilterView
            }
        }
    })

    if ($btnClearRowFilter) {
        $btnClearRowFilter.add_Click({
            $script:ActiveFilterPairId = $null
            & $FilterView
        })
    }

    $script:SearchTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:SearchTimer.Interval = [TimeSpan]::FromMilliseconds(600)
    $script:SearchTimer.add_Tick({ $script:SearchTimer.Stop(); & $FilterView })
    $txtSearch.add_TextChanged({
        if ($null -ne $script:ActiveFilterPairId) { $script:ActiveFilterPairId = $null }
        $script:SearchTimer.Stop()
        $script:SearchTimer.Start()
    })
    $cmbStatus.add_SelectionChanged({
        if ($null -ne $script:ActiveFilterPairId) { $script:ActiveFilterPairId = $null }
        & $FilterView
    })
    $btnReset.add_Click({
        $script:ActiveFilterPairId = $null
        $script:PreviousSelectedPairId = $null
        $cmbStatus.SelectedIndex = 0
        $txtSearch.Text = ''
        & $FilterView
    })

    # Sync selection
    $SyncSelection = {
        param($srcGrid, $targetGrid)
        if ($script:IgnoreSelectionChange -or -not $chkSync.IsChecked -or $null -ne $script:ActiveFilterPairId) { return }
        $script:IgnoreSelectionChange = $true
        try {
            $idx = $srcGrid.SelectedIndex
            if ($idx -ge 0 -and $idx -lt $targetGrid.Items.Count -and $targetGrid.SelectedIndex -ne $idx) {
                $targetGrid.SelectedIndex = $idx
                $targetGrid.ScrollIntoView($targetGrid.SelectedItem)
            }
        }
        finally {
            $script:IgnoreSelectionChange = $false
        }
    }
    $dgBase.add_SelectionChanged({ param($s, $e) & $SyncSelection $dgBase $dgUpdate })
    $dgUpdate.add_SelectionChanged({ param($s, $e) & $SyncSelection $dgUpdate $dgBase })

    # Sync scroll
    $script:SyncLock = $false
    $Window.add_Loaded({
            $script:svB = Get-ScrollViewer $dgBase
            $script:svU = Get-ScrollViewer $dgUpdate
            if ($script:svB -and $script:svU) {
                $script:svB.add_ScrollChanged({
                        param($s, $e)
                        if ($chkSync.IsChecked -and -not $script:SyncLock) {
                            $script:SyncLock = $true
                            if ($e.VerticalChange -ne 0) { $script:svU.ScrollToVerticalOffset($e.VerticalOffset) }
                            if ($e.HorizontalChange -ne 0) { $script:svU.ScrollToHorizontalOffset($e.HorizontalOffset) }
                            $script:SyncLock = $false
                        }
                    })
                $script:svU.add_ScrollChanged({
                        param($s, $e)
                        if ($chkSync.IsChecked -and -not $script:SyncLock) {
                            $script:SyncLock = $true
                            if ($e.VerticalChange -ne 0) { $script:svB.ScrollToVerticalOffset($e.VerticalOffset) }
                            if ($e.HorizontalChange -ne 0) { $script:svB.ScrollToHorizontalOffset($e.HorizontalOffset) }
                            $script:SyncLock = $false
                        }
                    })
            }
        })

    # Export
    $btnExport.add_Click({
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = 'Excel Files|*.xlsx'
            $sfd.FileName = "ExcelDiff_$(Get-Date -Format 'yyyy-MM-dd_HHmm').xlsx"
            $sfd.Title = Get-Loc 'DashExport'
            if ($sfd.ShowDialog() -eq 'OK') {
                try {
                    $curB = @($dgBase.ItemsSource)
                    $curU = @($dgUpdate.ItemsSource)
                    $exportRows = for ($i = 0; $i -lt $curB.Count; $i++) {
                        $b = $curB[$i]; $u = $curU[$i]
                        $row = [ordered]@{ _DiffStatus = $b._DiffStatus; _ChangedColumns = $b._ChangedColumns }
                        foreach ($p in $BaseAllProps) { $row["Base_$p"] = if ($null -ne $b.$p) { $b.$p } else { '' } }
                        foreach ($p in $UpdateAllProps) { $row["Update_$p"] = if ($null -ne $u.$p) { $u.$p } else { '' } }
                        [PSCustomObject]$row
                    }
                    $exportRows | Export-Excel -Path $sfd.FileName -AutoSize -FreezeTopRow -BoldTopRow
                    [System.Windows.MessageBox]::Show("$(Get-Loc 'ExportSaved')`n$($sfd.FileName)", (Get-Loc 'DialogInfo'))
                }
                catch {
                    [System.Windows.MessageBox]::Show("Export failed: $_", (Get-Loc 'DialogWarn'))
                }
            }
        })

    & $FilterView
    $Window.ShowDialog() | Out-Null
}
#endregion

function Show-CompareWizard {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Excel File Comparator"
        Height="720" Width="960"
        WindowStartupLocation="CenterScreen"
        Topmost="True"
        Background="{DynamicResource BgWindowBrush}">
    <Window.Resources>
        <SolidColorBrush x:Key="BgWindowBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgHeaderBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="BgCardBrush" Color="#1E293B"/>
        <SolidColorBrush x:Key="BgInputBrush" Color="#334155"/>
        <SolidColorBrush x:Key="BgButtonDefaultBrush" Color="#334155"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#475569"/>
        <SolidColorBrush x:Key="BorderSubtleBrush" Color="#334155"/>
        <SolidColorBrush x:Key="GridLinesBrush" Color="#334155"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#CBD5E1"/>
        <SolidColorBrush x:Key="TextMutedBrush" Color="#94A3B8"/>
        <SolidColorBrush x:Key="RowBgBrush" Color="#1E293B"/>
        <SolidColorBrush x:Key="RowAltBgBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="HeaderBgBrush" Color="#334155"/>
        <SolidColorBrush x:Key="HeaderFgBrush" Color="#F8FAFC"/>
        <SolidColorBrush x:Key="TabActiveBgBrush" Color="#1E293B"/>
        <SolidColorBrush x:Key="TabInactiveBgBrush" Color="#0F172A"/>
        <SolidColorBrush x:Key="TabHoverBgBrush" Color="#334155"/>
        <SolidColorBrush x:Key="ItemHoverBgBrush" Color="#334155"/>

        <!-- TabControl Style -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TabPanel Grid.Row="0" IsItemsHost="True" Margin="2,0,2,0"/>
                            <Border Grid.Row="1"
                                    Background="{DynamicResource BgCardBrush}"
                                    BorderBrush="{DynamicResource BorderBrush}"
                                    BorderThickness="1"
                                    CornerRadius="0,6,6,6">
                                <ContentPresenter ContentSource="SelectedContent" Margin="0"/>
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TabItem Style -->
        <Style TargetType="TabItem">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="{DynamicResource TextSecondaryBrush}"/>
            <Setter Property="Padding" Value="18,10"/>
            <Setter Property="Margin" Value="0,0,4,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder"
                                CornerRadius="6,6,0,0"
                                Padding="{TemplateBinding Padding}"
                                Margin="{TemplateBinding Margin}"
                                Background="{DynamicResource TabInactiveBgBrush}"
                                BorderBrush="{DynamicResource BorderBrush}"
                                BorderThickness="1,1,1,0">
                            <ContentPresenter x:Name="TabContent"
                                              ContentSource="Header"
                                              HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              TextElement.Foreground="{DynamicResource TextSecondaryBrush}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource BgCardBrush}"/>
                                <Setter TargetName="TabBorder" Property="BorderThickness" Value="1,2,1,0"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#2563EB"/>
                                <Setter TargetName="TabContent" Property="TextElement.Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
                                <Setter TargetName="TabContent" Property="TextElement.FontWeight" Value="Bold"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{DynamicResource TabHoverBgBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ComboBox Style -->
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource BgInputBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton"
                                          Focusable="false"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}">
                                <ToggleButton.Template>
                                    <ControlTemplate TargetType="ToggleButton">
                                        <Border x:Name="Border" CornerRadius="4" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                                            <Path x:Name="Arrow" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" Fill="{DynamicResource TextSecondaryBrush}" Data="M 0 0 L 4 4 L 8 0 Z"/>
                                        </Border>
                                    </ControlTemplate>
                                </ToggleButton.Template>
                            </ToggleButton>
                            <ContentPresenter x:Name="ContentSite"
                                              IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="8,4,24,4"
                                              VerticalAlignment="Center"
                                              HorizontalAlignment="Left"
                                              TextElement.Foreground="{TemplateBinding Foreground}"/>
                            <Popup x:Name="Popup"
                                   Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True"
                                   Focusable="False"
                                   PopupAnimation="Slide">
                                <Border x:Name="DropDownBorder"
                                        Background="{DynamicResource BgCardBrush}"
                                        BorderBrush="{DynamicResource BorderBrush}"
                                        BorderThickness="1"
                                        CornerRadius="4"
                                        MinWidth="{TemplateBinding ActualWidth}"
                                        MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <ScrollViewer Margin="2" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained" />
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Border" Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource ItemHoverBgBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- DataGrid Style -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource BgCardBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="RowBackground" Value="{DynamicResource RowBgBrush}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource RowAltBgBrush}"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource GridLinesBrush}"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Background" Value="Transparent"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#2563EB"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
        </Style>

        <!-- ListBoxItem Style -->
        <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="Border" Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#2563EB"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{DynamicResource ItemHoverBgBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border x:Name="borderWizHeader" Grid.Row="0" Background="{DynamicResource BgHeaderBrush}" Padding="16,12">
            <DockPanel>
                <TextBlock x:Name="lblWizTitle" Foreground="White" FontSize="16" FontWeight="Bold" VerticalAlignment="Center"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnWizTheme" Height="28" Padding="12,0" Margin="0,0,14,0"
                            Background="{DynamicResource BgInputBrush}" Foreground="{DynamicResource TextPrimaryBrush}"
                            BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1"
                            FontWeight="SemiBold" FontSize="12" Cursor="Hand"/>
                    <TextBlock x:Name="lblLang" Foreground="#D1D5DB" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <ComboBox x:Name="cmbLang" Width="65" SelectedIndex="1">
                        <ComboBoxItem Content="EN"/>
                        <ComboBoxItem Content="PL"/>
                        <ComboBoxItem Content="DE"/>
                    </ComboBox>
                </StackPanel>
            </DockPanel>
        </Border>

        <TabControl Grid.Row="1" x:Name="tabs" Margin="14,10,14,10">

            <TabItem x:Name="tabFiles" Header="Files">
                <StackPanel Margin="20">
                    <Border x:Name="cardBase" Background="{DynamicResource BgInputBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="16" Margin="0,0,0,16">
                        <StackPanel>
                            <TextBlock x:Name="lblBaseFile" FontWeight="Bold" FontSize="14" Foreground="{DynamicResource TextPrimaryBrush}" Margin="0,0,0,12"/>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="btnBaseFile" Width="180" Height="34" Background="#2563EB" Foreground="White"
                                        FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                                <TextBlock x:Name="txtBasePath" Text="---" VerticalAlignment="Center"
                                           Margin="14,0,20,0" Foreground="{DynamicResource TextMutedBrush}" Width="310"
                                           TextTrimming="CharacterEllipsis"/>
                                <TextBlock x:Name="lblSheet1" Foreground="{DynamicResource TextPrimaryBrush}" VerticalAlignment="Center" Margin="0,0,8,0" FontWeight="SemiBold"/>
                                <ComboBox x:Name="cmbBaseSheet" Width="170" Height="30"/>
                            </StackPanel>
                        </StackPanel>
                    </Border>

                    <Border x:Name="cardUpdate" Background="{DynamicResource BgInputBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="16">
                        <StackPanel>
                            <TextBlock x:Name="lblUpdateFile" FontWeight="Bold" FontSize="14" Foreground="#14B8A6" Margin="0,0,0,12"/>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="btnUpdateFile" Width="180" Height="34" Background="#0D9488" Foreground="White"
                                        FontWeight="Bold" BorderThickness="0" Cursor="Hand"/>
                                <TextBlock x:Name="txtUpdatePath" Text="---" VerticalAlignment="Center"
                                           Margin="14,0,20,0" Foreground="{DynamicResource TextMutedBrush}" Width="310"
                                           TextTrimming="CharacterEllipsis"/>
                                <TextBlock x:Name="lblSheet2" Foreground="{DynamicResource TextPrimaryBrush}" VerticalAlignment="Center" Margin="0,0,8,0" FontWeight="SemiBold"/>
                                <ComboBox x:Name="cmbUpdateSheet" Width="170" Height="30"/>
                            </StackPanel>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </TabItem>

            <TabItem x:Name="tabMapping" Header="Column Mapping">
                <Grid Margin="16">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                        <Button x:Name="btnAutoMap"   Width="200" Height="32" Margin="0,0,10,0"
                                Background="#8B5CF6" Foreground="White" FontWeight="Bold"
                                BorderThickness="0" Cursor="Hand"/>
                        <Button x:Name="btnAddMap"    Width="170" Height="32" Margin="0,0,10,0"
                                Background="#10B981" Foreground="White" FontWeight="Bold"
                                BorderThickness="0" Cursor="Hand"/>
                        <Button x:Name="btnEditMap"   Width="140" Height="32" Margin="0,0,10,0"
                                Background="#2563EB" Foreground="White" FontWeight="Bold"
                                BorderThickness="0" Cursor="Hand"/>
                        <Button x:Name="btnRemoveMap" Width="140" Height="32" Margin="0,0,10,0"
                                Background="#EF4444" Foreground="White" FontWeight="Bold"
                                BorderThickness="0" Cursor="Hand"/>
                        <Button x:Name="btnClearMap"  Width="120" Height="32"
                                Background="#64748B" Foreground="White" FontWeight="Bold"
                                BorderThickness="0" Cursor="Hand"/>
                    </StackPanel>
                    <DataGrid x:Name="dgMappings" Grid.Row="1"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              CanUserSortColumns="False" SelectionMode="Single"
                              GridLinesVisibility="Horizontal" HeadersVisibility="Column"
                              BorderThickness="1" Margin="0,0,0,8">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="FontWeight" Value="Bold"/>
                                <Setter Property="Background" Value="{DynamicResource HeaderBgBrush}"/>
                                <Setter Property="Foreground" Value="{DynamicResource HeaderFgBrush}"/>
                                <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
                                <Setter Property="BorderThickness" Value="0,0,1,1"/>
                                <Setter Property="Padding"    Value="8,5"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn x:Name="dgcBase"   Binding="{Binding BaseColsDisplay}"   Width="*">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
                                        <Setter Property="Padding" Value="4,2"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn x:Name="dgcMerge"  Binding="{Binding MergeMode}"         Width="120">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
                                        <Setter Property="Padding" Value="4,2"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn x:Name="dgcSep"    Binding="{Binding Separator}"         Width="70">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
                                        <Setter Property="Padding" Value="4,2"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn x:Name="dgcUpdate" Binding="{Binding UpdateColsDisplay}" Width="*">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
                                        <Setter Property="Padding" Value="4,2"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn x:Name="dgcLabel"  Binding="{Binding Label}"             Width="*">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}"/>
                                        <Setter Property="Padding" Value="4,2"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                    <TextBlock Grid.Row="2" Text="Tip: Use Ctrl+Click to select multiple columns in the mapping dialog."
                               Foreground="{DynamicResource TextMutedBrush}" FontSize="11" FontStyle="Italic"/>
                </Grid>
            </TabItem>

            <TabItem x:Name="tabJoinKey" Header="Join Key">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{DynamicResource BgCardBrush}" CornerRadius="6" Padding="14,10" Margin="0,0,0,14" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1">
                        <StackPanel>
                            <RadioButton x:Name="rbJoinKeyCols" Content="Match rows using Key Column(s)" IsChecked="True" GroupName="JoinModeGroup" Foreground="{DynamicResource TextPrimaryBrush}" FontWeight="SemiBold" FontSize="13" Margin="0,0,0,6"/>
                            <RadioButton x:Name="rbJoinByRow"   Content="Match rows by Row Order (Sequential / Line-by-Line)" GroupName="JoinModeGroup" Foreground="{DynamicResource TextPrimaryBrush}" FontWeight="SemiBold" FontSize="13"/>
                            <TextBlock x:Name="lblJoinRowHint" Text="Compares Row 1 to Row 1, Row 2 to Row 2 sequentially. No key columns required." Foreground="{DynamicResource TextMutedBrush}" FontSize="11" Margin="22,4,0,0" FontStyle="Italic"/>
                        </StackPanel>
                    </Border>
                    <TextBlock x:Name="lblJoinKey" Grid.Row="1" Foreground="{DynamicResource TextPrimaryBrush}" FontWeight="SemiBold" FontSize="13" Margin="0,0,0,12" TextWrapping="Wrap"/>
                    <Grid x:Name="gridJoinCols" Grid.Row="2">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,10,0">
                            <TextBlock x:Name="lblJoinBase" Foreground="{DynamicResource TextPrimaryBrush}" FontWeight="Bold" Margin="0,0,0,8"/>
                            <ListBox x:Name="lbJoinBase" SelectionMode="Extended" Background="{DynamicResource BgInputBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="10,0,0,0">
                            <TextBlock x:Name="lblJoinUpdate" Foreground="{DynamicResource TextPrimaryBrush}" FontWeight="Bold" Margin="0,0,0,8"/>
                            <ListBox x:Name="lbJoinUpdate" SelectionMode="Extended" Background="{DynamicResource BgInputBrush}" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="1"/>
                        </StackPanel>
                    </Grid>
                </Grid>
            </TabItem>

            <TabItem x:Name="tabOptions" Header="Options">
                <StackPanel Margin="24">
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                        <TextBlock x:Name="lblOptLayout" Text="Layout:" Foreground="{DynamicResource TextPrimaryBrush}" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="SemiBold" FontSize="13"/>
                        <ComboBox x:Name="cmbOptLayout" Width="220" Height="28" SelectedIndex="0">
                            <ComboBoxItem x:Name="optVert" Content="Top / Bottom (Vertical)"/>
                            <ComboBoxItem x:Name="optHoriz" Content="Side by Side (Horizontal)"/>
                        </ComboBox>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                        <TextBlock x:Name="lblOptTheme" Text="Theme:" Foreground="{DynamicResource TextPrimaryBrush}" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="SemiBold" FontSize="13"/>
                        <ComboBox x:Name="cmbOptTheme" Width="220" Height="28" SelectedIndex="0">
                            <ComboBoxItem x:Name="optThemeLight" Content="☀️ Light"/>
                            <ComboBoxItem x:Name="optThemeDark"  Content="🌙 Dark"/>
                        </ComboBox>
                    </StackPanel>
                    <CheckBox x:Name="chkIgnoreCase"         Foreground="{DynamicResource TextPrimaryBrush}" IsChecked="True"  FontSize="13" Margin="0,0,0,12"/>
                    <CheckBox x:Name="chkTrim"               Foreground="{DynamicResource TextPrimaryBrush}" IsChecked="True"  FontSize="13" Margin="0,0,0,12"/>
                    <CheckBox x:Name="chkIgnoreSpecialChars" Foreground="{DynamicResource TextPrimaryBrush}" IsChecked="True"  FontSize="13" Margin="0,0,0,12"/>
                    <CheckBox x:Name="chkIgnoreAllSpaces"    Foreground="{DynamicResource TextPrimaryBrush}" IsChecked="True"  FontSize="13" Margin="0,0,0,12"/>
                    <CheckBox x:Name="chkHideUnchanged"      Foreground="{DynamicResource TextPrimaryBrush}" IsChecked="False" FontSize="13"/>
                </StackPanel>
            </TabItem>
        </TabControl>

        <Border x:Name="borderBottom" Grid.Row="2" Background="{DynamicResource BgCardBrush}" Padding="16,10" BorderBrush="{DynamicResource BorderBrush}" BorderThickness="0,1,0,0">
            <Button x:Name="btnRun" Height="46" FontSize="17" FontWeight="Bold"
                    Background="#10B981" Foreground="White" BorderThickness="0" Cursor="Hand"/>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $Wiz = [Windows.Markup.XamlReader]::Load($reader)

    $borderWizHeader = $Wiz.FindName('borderWizHeader')
    $borderBottom = $Wiz.FindName('borderBottom')
    $cardBase = $Wiz.FindName('cardBase')
    $cardUpdate = $Wiz.FindName('cardUpdate')
    $cmbLang = $Wiz.FindName('cmbLang')
    $lblWizTitle = $Wiz.FindName('lblWizTitle')
    $lblLang = $Wiz.FindName('lblLang')
    $btnWizTheme = $Wiz.FindName('btnWizTheme')
    $lblBaseFile = $Wiz.FindName('lblBaseFile')
    $btnBaseFile = $Wiz.FindName('btnBaseFile')
    $txtBasePath = $Wiz.FindName('txtBasePath')
    $cmbBaseSheet = $Wiz.FindName('cmbBaseSheet')
    $lblUpdateFile = $Wiz.FindName('lblUpdateFile')
    $btnUpdateFile = $Wiz.FindName('btnUpdateFile')
    $txtUpdatePath = $Wiz.FindName('txtUpdatePath')
    $cmbUpdateSheet = $Wiz.FindName('cmbUpdateSheet')
    $dgMappings = $Wiz.FindName('dgMappings')
    $btnAutoMap = $Wiz.FindName('btnAutoMap')
    $btnAddMap = $Wiz.FindName('btnAddMap')
    $btnEditMap = $Wiz.FindName('btnEditMap')
    $btnRemoveMap = $Wiz.FindName('btnRemoveMap')
    $btnClearMap = $Wiz.FindName('btnClearMap')
    $rbJoinKeyCols = $Wiz.FindName('rbJoinKeyCols')
    $rbJoinByRow = $Wiz.FindName('rbJoinByRow')
    $lblJoinRowHint = $Wiz.FindName('lblJoinRowHint')
    $gridJoinCols = $Wiz.FindName('gridJoinCols')
    $lblJoinKey = $Wiz.FindName('lblJoinKey')
    $lblJoinBase = $Wiz.FindName('lblJoinBase')
    $lblJoinUpdate = $Wiz.FindName('lblJoinUpdate')
    $lbJoinBase = $Wiz.FindName('lbJoinBase')
    $lbJoinUpdate = $Wiz.FindName('lbJoinUpdate')
    $lblOptLayout = $Wiz.FindName('lblOptLayout')
    $cmbOptLayout = $Wiz.FindName('cmbOptLayout')
    $optVert = $Wiz.FindName('optVert')
    $optHoriz = $Wiz.FindName('optHoriz')
    $lblOptTheme = $Wiz.FindName('lblOptTheme')
    $cmbOptTheme = $Wiz.FindName('cmbOptTheme')
    $optThemeLight = $Wiz.FindName('optThemeLight')
    $optThemeDark = $Wiz.FindName('optThemeDark')
    $chkIgnoreCase = $Wiz.FindName('chkIgnoreCase')
    $chkTrim = $Wiz.FindName('chkTrim')
    $chkIgnoreSpecialChars = $Wiz.FindName('chkIgnoreSpecialChars')
    $chkIgnoreAllSpaces = $Wiz.FindName('chkIgnoreAllSpaces')
    $chkHideUnchanged = $Wiz.FindName('chkHideUnchanged')
    $btnRun = $Wiz.FindName('btnRun')

    $UpdateJoinModeUI = {
        if ($rbJoinByRow.IsChecked) {
            $gridJoinCols.IsEnabled = $false
            $gridJoinCols.Opacity = 0.35
            $lblJoinKey.Text = Get-Loc 'LblRowIndexJoin'
        }
        else {
            $gridJoinCols.IsEnabled = $true
            $gridJoinCols.Opacity = 1.0
            $lblJoinKey.Text = Get-Loc 'LblJoinKey'
        }
    }
    if ($rbJoinKeyCols) { $rbJoinKeyCols.add_Checked($UpdateJoinModeUI) }
    if ($rbJoinByRow)   { $rbJoinByRow.add_Checked($UpdateJoinModeUI) }

    $ApplyLocalization = {
        $lblWizTitle.Text = Get-Loc 'WizardTitle'
        $Wiz.Title = Get-Loc 'WizardTitle'
        $lblLang.Text = Get-Loc 'Language'
        if ($lblBaseFile) { $lblBaseFile.Text = Get-Loc 'LblBaseFile' }
        if ($lblUpdateFile) { $lblUpdateFile.Text = Get-Loc 'LblUpdateFile' }
        $btnBaseFile.Content = Get-Loc 'BtnBaseFile'
        $btnUpdateFile.Content = Get-Loc 'BtnUpdateFile'
        $Wiz.FindName('lblSheet1').Text = Get-Loc 'SheetLabel'
        $Wiz.FindName('lblSheet2').Text = Get-Loc 'SheetLabel'
        if ($btnAutoMap) { $btnAutoMap.Content = Get-Loc 'BtnAutoMap' }
        $btnAddMap.Content = Get-Loc 'BtnAddMapping'
        $btnEditMap.Content = Get-Loc 'BtnEditMapping'
        $btnRemoveMap.Content = Get-Loc 'BtnRemoveMapping'
        if ($btnClearMap) { $btnClearMap.Content = Get-Loc 'BtnClearMap' }
        $Wiz.FindName('dgcBase').Header = Get-Loc 'ColBaseColumns'
        $Wiz.FindName('dgcMerge').Header = Get-Loc 'ColMergeMode'
        $Wiz.FindName('dgcSep').Header = Get-Loc 'ColSeparator'
        $Wiz.FindName('dgcUpdate').Header = Get-Loc 'ColUpdateColumns'
        $Wiz.FindName('dgcLabel').Header = Get-Loc 'ColLabel'
        if ($rbJoinKeyCols) { $rbJoinKeyCols.Content = Get-Loc 'JoinModeKey' }
        if ($rbJoinByRow)   { $rbJoinByRow.Content = Get-Loc 'JoinModeRow' }
        if ($lblJoinRowHint) { $lblJoinRowHint.Text = Get-Loc 'TipMatchByRow' }
        if ($rbJoinByRow -and $rbJoinByRow.IsChecked) {
            $lblJoinKey.Text = Get-Loc 'LblRowIndexJoin'
        }
        else {
            $lblJoinKey.Text = Get-Loc 'LblJoinKey'
        }
        $lblJoinBase.Text = Get-Loc 'LblJoinBase'
        $lblJoinUpdate.Text = Get-Loc 'LblJoinUpdate'
        if ($lblOptLayout) { $lblOptLayout.Text = Get-Loc 'LayoutMode' }
        if ($optVert) { $optVert.Content = Get-Loc 'LayoutVert' }
        if ($optHoriz) { $optHoriz.Content = Get-Loc 'LayoutHoriz' }
        if ($lblOptTheme) { $lblOptTheme.Text = Get-Loc 'ThemeLabel' }
        if ($optThemeLight) { $optThemeLight.Content = Get-Loc 'ThemeLight' }
        if ($optThemeDark) { $optThemeDark.Content = Get-Loc 'ThemeDark' }
        $chkIgnoreCase.Content = Get-Loc 'LblIgnoreCase'
        $chkTrim.Content = Get-Loc 'LblTrimWhitespace'
        if ($chkIgnoreSpecialChars) { $chkIgnoreSpecialChars.Content = Get-Loc 'LblIgnoreSpecialChars' }
        if ($chkIgnoreAllSpaces) { $chkIgnoreAllSpaces.Content = Get-Loc 'LblIgnoreAllSpaces' }
        $chkHideUnchanged.Content = Get-Loc 'LblIgnoreUnchanged'
        $btnRun.Content = Get-Loc 'BtnRunCompare'
        $Wiz.FindName('tabFiles').Header = Get-Loc 'TabFiles'
        $Wiz.FindName('tabMapping').Header = Get-Loc 'TabMapping'
        $Wiz.FindName('tabJoinKey').Header = Get-Loc 'TabJoinKey'
        $Wiz.FindName('tabOptions').Header = Get-Loc 'TabOptions'
    }

    $ApplyWizardTheme = {
        $pal = Get-ThemePalette $Global:CurrentTheme
        Update-WpfThemeResources $Wiz $pal

        if ($btnWizTheme) {
            $btnWizTheme.Content = if ($pal.IsDark) { Get-Loc 'ThemeLight' } else { Get-Loc 'ThemeDark' }
        }

        if ($cmbOptTheme) {
            $cmbOptTheme.SelectedIndex = if ($pal.IsDark) { 1 } else { 0 }
        }

        Set-WindowDarkMode -Window $Wiz -IsDark $pal.IsDark
    }

    $script:BasePath = ''
    $script:UpdatePath = ''
    $script:BaseHdrs = @()
    $script:UpdateHdrs = @()
    $script:MapRules = [System.Collections.ArrayList]::new()
    Load-AppSettings
    $li = @('EN', 'PL', 'DE').IndexOf($Global:CurrentLang)
    if ($li -ge 0) { $cmbLang.SelectedIndex = $li }
    & $ApplyLocalization
    & $ApplyWizardTheme

    $Wiz.add_Loaded({
            & $ApplyWizardTheme
        })

    $btnWizTheme.add_Click({
            $Global:CurrentTheme = if ($Global:CurrentTheme -eq 'Dark') { 'Light' } else { 'Dark' }
            try { Save-AppSettings } catch {}
            & $ApplyWizardTheme
        })

    $cmbOptTheme.add_SelectionChanged({
            if ($cmbOptTheme.SelectedItem) {
                $newTheme = if ($cmbOptTheme.SelectedIndex -eq 1) { 'Dark' } else { 'Light' }
                if ($newTheme -ne $Global:CurrentTheme) {
                    $Global:CurrentTheme = $newTheme
                    try { Save-AppSettings } catch {}
                    & $ApplyWizardTheme
                }
            }
        })

    $cmbLang.add_SelectionChanged({
            $Global:CurrentLang = $cmbLang.SelectedItem.Content.ToString()
            try { Save-AppSettings } catch {}
            & $ApplyLocalization
            & $ApplyWizardTheme
        })

    $RefreshJoin = {
        $lbJoinBase.Items.Clear(); $lbJoinUpdate.Items.Clear()
        foreach ($h in $script:BaseHdrs) { [void]$lbJoinBase.Items.Add($h) }
        foreach ($h in $script:UpdateHdrs) { [void]$lbJoinUpdate.Items.Add($h) }
    }

    $RefreshMapGrid = {
        $rows = $script:MapRules | ForEach-Object {
            [PSCustomObject]@{
                BaseColsDisplay   = $_.BaseColumns -join ' | '
                MergeMode         = $_.MergeMode
                Separator         = $_.Separator
                UpdateColsDisplay = $_.UpdateColumns -join ' | '
                Label             = $_.Label
                _Rule             = $_
            }
        }
        $dgMappings.ItemsSource = [System.Collections.ArrayList]@($rows)
    }

    $AutoMapCols = {
        param([bool]$Silent = $false)
        if ($script:BaseHdrs.Count -eq 0 -or $script:UpdateHdrs.Count -eq 0) {
            if (-not $Silent) {
                [System.Windows.MessageBox]::Show((Get-Loc 'WarnNoBase'), (Get-Loc 'DialogWarn'))
            }
            return
        }

        $alreadyMappedBase = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($r in $script:MapRules) {
            foreach ($c in $r.BaseColumns) { [void]$alreadyMappedBase.Add($c) }
        }

        $updLookup = @{}
        foreach ($u in $script:UpdateHdrs) {
            $key = $u.Trim().ToLowerInvariant()
            if (-not $updLookup.ContainsKey($key)) {
                $updLookup[$key] = $u
            }
        }

        $addedCount = 0
        foreach ($b in $script:BaseHdrs) {
            if ($alreadyMappedBase.Contains($b)) { continue }
            $bKey = $b.Trim().ToLowerInvariant()
            if ($updLookup.ContainsKey($bKey)) {
                $uMatch = $updLookup[$bKey]
                $rule = @{
                    BaseColumns   = @($b)
                    UpdateColumns = @($uMatch)
                    MergeMode     = 'Exact'
                    Separator     = ''
                    Label         = $b
                }
                [void]$script:MapRules.Add($rule)
                [void]$alreadyMappedBase.Add($b)
                $addedCount++
            }
        }

        if ($addedCount -gt 0) {
            & $RefreshMapGrid
        }

        # Auto-select join key if none selected
        if ($lbJoinBase.SelectedItems.Count -eq 0 -and $lbJoinUpdate.SelectedItems.Count -eq 0) {
            $preferredJoinKeys = @('id', 'lp', 'kod', 'nr', 'key', 'imię i nazwisko dziecka', 'imie i nazwisko dziecka')
            $foundJoin = $false
            foreach ($pk in $preferredJoinKeys) {
                $bIdx = -1; $uIdx = -1
                for ($i = 0; $i -lt $lbJoinBase.Items.Count; $i++) {
                    if ($lbJoinBase.Items[$i].ToString().Trim().ToLowerInvariant() -eq $pk) { $bIdx = $i; break }
                }
                for ($i = 0; $i -lt $lbJoinUpdate.Items.Count; $i++) {
                    if ($lbJoinUpdate.Items[$i].ToString().Trim().ToLowerInvariant() -eq $pk) { $uIdx = $i; break }
                }
                if ($bIdx -ge 0 -and $uIdx -ge 0) {
                    $lbJoinBase.SelectedIndex = $bIdx
                    $lbJoinUpdate.SelectedIndex = $uIdx
                    $foundJoin = $true
                    break
                }
            }
            if (-not $foundJoin -and $lbJoinBase.Items.Count -gt 0 -and $lbJoinUpdate.Items.Count -gt 0) {
                for ($i = 0; $i -lt $lbJoinBase.Items.Count; $i++) {
                    $bName = $lbJoinBase.Items[$i].ToString().Trim().ToLowerInvariant()
                    for ($j = 0; $j -lt $lbJoinUpdate.Items.Count; $j++) {
                        $uName = $lbJoinUpdate.Items[$j].ToString().Trim().ToLowerInvariant()
                        if ($bName -eq $uName) {
                            $lbJoinBase.SelectedIndex = $i
                            $lbJoinUpdate.SelectedIndex = $j
                            $foundJoin = $true
                            break
                        }
                    }
                    if ($foundJoin) { break }
                }
            }
        }

        if (-not $Silent) {
            if ($addedCount -gt 0) {
                $msg = [string]::Format((Get-Loc 'MsgAutoMapped'), $addedCount)
                [System.Windows.MessageBox]::Show($msg, (Get-Loc 'DialogInfo'))
            }
            else {
                [System.Windows.MessageBox]::Show((Get-Loc 'WarnNoMatchingCols'), (Get-Loc 'DialogInfo'))
            }
        }
    }

    $btnBaseFile.add_Click({
            $p = Pick-File $txtBasePath $cmbBaseSheet; if ($p) { $script:BasePath = $p }
        })
    $btnUpdateFile.add_Click({
            $p = Pick-File $txtUpdatePath $cmbUpdateSheet; if ($p) { $script:UpdatePath = $p }
        })
    $cmbBaseSheet.add_SelectionChanged({
            if ($txtBasePath.Text -ne '---' -and $cmbBaseSheet.SelectedItem) {
                $script:BaseHdrs = Get-ExcelHeaders -Path $txtBasePath.Text -Sheet $cmbBaseSheet.SelectedItem
                & $RefreshJoin
                if ($script:BaseHdrs.Count -gt 0 -and $script:UpdateHdrs.Count -gt 0 -and $script:MapRules.Count -eq 0) {
                    & $AutoMapCols $true
                }
            }
        })
    $cmbUpdateSheet.add_SelectionChanged({
            if ($txtUpdatePath.Text -ne '---' -and $cmbUpdateSheet.SelectedItem) {
                $script:UpdateHdrs = Get-ExcelHeaders -Path $txtUpdatePath.Text -Sheet $cmbUpdateSheet.SelectedItem
                & $RefreshJoin
                if ($script:BaseHdrs.Count -gt 0 -and $script:UpdateHdrs.Count -gt 0 -and $script:MapRules.Count -eq 0) {
                    & $AutoMapCols $true
                }
            }
        })

    $btnAutoMap.add_Click({
            & $AutoMapCols $false
        })
    $btnClearMap.add_Click({
            $script:MapRules.Clear()
            & $RefreshMapGrid
        })
    $btnAddMap.add_Click({
            if ($script:BaseHdrs.Count -eq 0 -or $script:UpdateHdrs.Count -eq 0) {
                [System.Windows.MessageBox]::Show((Get-Loc 'WarnNoBase'), (Get-Loc 'DialogWarn')); return
            }
            $rule = Show-MappingRuleDialog -BaseHeaders $script:BaseHdrs -UpdateHeaders $script:UpdateHdrs
            if ($null -ne $rule) { [void]$script:MapRules.Add($rule); & $RefreshMapGrid }
        })
    $btnEditMap.add_Click({
            $sel = $dgMappings.SelectedItem; if ($null -eq $sel) { return }
            $existing = $sel._Rule
            $idx = $script:MapRules.IndexOf($existing)
            $newRule = Show-MappingRuleDialog -BaseHeaders $script:BaseHdrs -UpdateHeaders $script:UpdateHdrs -ExistingRule $existing
            if ($null -ne $newRule -and $idx -ge 0) { $script:MapRules[$idx] = $newRule; & $RefreshMapGrid }
        })
    $btnRemoveMap.add_Click({
            $sel = $dgMappings.SelectedItem
            if ($null -ne $sel) { $script:MapRules.Remove($sel._Rule); & $RefreshMapGrid }
        })

    $btnRun.add_Click({
            if ([string]::IsNullOrWhiteSpace($script:BasePath)) {
                [System.Windows.MessageBox]::Show((Get-Loc 'WarnNoBase'), (Get-Loc 'DialogWarn')); return
            }
            if ([string]::IsNullOrWhiteSpace($script:UpdatePath)) {
                [System.Windows.MessageBox]::Show((Get-Loc 'WarnNoUpdate'), (Get-Loc 'DialogWarn')); return
            }
            if ($script:MapRules.Count -eq 0) {
                [System.Windows.MessageBox]::Show((Get-Loc 'WarnNoMapping'), (Get-Loc 'DialogWarn')); return
            }

            $matchByRow = ($null -ne $rbJoinByRow -and [bool]$rbJoinByRow.IsChecked)
            $joinBase = @()
            $joinUpdate = @()
            if (-not $matchByRow) {
                $joinBase = @($lbJoinBase.SelectedItems)
                $joinUpdate = @($lbJoinUpdate.SelectedItems)
                if ($joinBase.Count -eq 0) {
                    [System.Windows.MessageBox]::Show((Get-Loc 'WarnNoJoinKey'), (Get-Loc 'DialogWarn')); return
                }
                if ($joinBase.Count -ne $joinUpdate.Count) {
                    [System.Windows.MessageBox]::Show((Get-Loc 'WarnJoinMismatch'), (Get-Loc 'DialogWarn')); return
                }
            }

            $loadFrm = New-Object System.Windows.Forms.Form
            $loadFrm.Text = Get-Loc 'WizardTitle'; $loadFrm.Size = New-Object System.Drawing.Size(340, 100)
            $loadFrm.StartPosition = 'CenterScreen'; $loadFrm.TopMost = $true
            $ll = New-Object System.Windows.Forms.Label; $ll.Text = Get-Loc 'LoadingData'
            $ll.AutoSize = $true; $ll.Location = New-Object System.Drawing.Point(40, 30)
            $loadFrm.Controls.Add($ll); $loadFrm.Show()
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $baseData = Load-ExcelData -Path $script:BasePath   -Sheet $cmbBaseSheet.SelectedItem
                $updateData = Load-ExcelData -Path $script:UpdatePath -Sheet $cmbUpdateSheet.SelectedItem

                if ($baseData.Count -eq 0) { throw 'Base file is empty or could not be read.' }
                if ($updateData.Count -eq 0) { throw 'Update file is empty or could not be read.' }

                $bProps = $baseData[0].PSObject.Properties.Name   | Where-Object { $_ -notmatch '^(RowError|RowState|Table|ItemArray|HasErrors)$' }
                $uProps = $updateData[0].PSObject.Properties.Name | Where-Object { $_ -notmatch '^(RowError|RowState|Table|ItemArray|HasErrors)$' }

                $diff = Invoke-ExcelDiff `
                    -BaseData           $baseData `
                    -UpdateData         $updateData `
                    -BaseAllProps       $bProps `
                    -UpdateAllProps     $uProps `
                    -MappingRules       @($script:MapRules) `
                    -JoinBaseColumns    $joinBase `
                    -JoinUpdateColumns  $joinUpdate `
                    -IgnoreCase         ([bool]$chkIgnoreCase.IsChecked) `
                    -TrimWhitespace     ([bool]$chkTrim.IsChecked) `
                    -IgnoreSpecialChars ([bool]$chkIgnoreSpecialChars.IsChecked) `
                    -IgnoreAllSpaces    ([bool]$chkIgnoreAllSpaces.IsChecked) `
                    -MatchByRowOrder    $matchByRow

                if ([bool]$chkHideUnchanged.IsChecked -and $diff.Base.Count -gt 0) {
                    $unch = Get-Loc 'StatusUnchanged'
                    $keepIdx = @(for ($i = 0; $i -lt $diff.Base.Count; $i++) {
                            if ($diff.Base[$i]._DiffStatus -ne $unch) { $i }
                        })
                    $diff = @{
                        Base   = [System.Collections.ArrayList]@($keepIdx | ForEach-Object { $diff.Base[$_] })
                        Update = [System.Collections.ArrayList]@($keepIdx | ForEach-Object { $diff.Update[$_] })
                    }
                }

                $loadFrm.Close()

                $chosenLayout = if ($cmbOptLayout.SelectedIndex -eq 1) { 'Horizontal' } else { 'Vertical' }

                Show-CompareResult `
                    -DiffBase       $diff.Base `
                    -DiffUpdate     $diff.Update `
                    -BaseAllProps   $bProps `
                    -UpdateAllProps $uProps `
                    -BaseFileName   (Split-Path $script:BasePath   -Leaf) `
                    -UpdateFileName (Split-Path $script:UpdatePath -Leaf) `
                    -MappingRules   @($script:MapRules) `
                    -DefaultLayout  $chosenLayout

            }
            catch {
                try { $loadFrm.Close() } catch {}
                [System.Windows.MessageBox]::Show("Error: $_`n`n$($_.ScriptStackTrace)", (Get-Loc 'DialogWarn'))
            }
        })

    $Wiz.ShowDialog() | Out-Null
}
#endregion

#region 7. Entry Point
Show-CompareWizard
#endregion
