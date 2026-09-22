Attribute VB_Name = "modVersion"
Option Explicit

' Version identity, and the About block on the Settings sheet.
'
' There are TWO version numbers and they mean different things:
'
'   APP_VERSION   this build of the workbook - code, layout, formulas.
'                 Changes every time something is rebuilt.
'
'   SCHEMA_VER    the shape of the stored data, in modUtils. Changes only
'                 when the job table's columns change, because that is what
'                 decides whether one workbook's records can be read by
'                 another (spec 11 collation, and the S_SchemaVer stamped on
'                 every job row).
'
' A build can go 0.4.0 -> 0.5.0 with the schema staying at 1.0, and that is
' the normal case. Conflating the two would force a schema bump on every
' cosmetic change and make the collation check meaningless.
'
' APP_VERSION is 0.<phase>.<revision>, tracking the build sequence in the
' design document: 0.7.x is "phases 1-7 built", and it reaches 1.0.0 when
' acceptance passes at the end of the sequence.

Public Const APP_NAME As String = "Print Cost Management"
' 0.9.0 - phase 9 begins: the 2026-09-22 post-phase-8 snag list. Real new
' functionality, not just fixes, so the phase digit moves again:
'
'   - Printer and Paper Stock on a location sheet now filter each other in
'     BOTH directions, auto-filling either one when only a single compatible
'     option remains. Previously Paper Stock was locked until a printer was
'     chosen; both cells now start unlocked and fully populated.
'     modCatalog.PrintersForStock is new (the reverse of the existing
'     StocksFor); modLists.BindPrinterRange/BindStockRange/AutoFillIfSingle
'     replace the old one-directional BindStockRange; modValidation's
'     OnPrinterChanged/OnStockChanged are now symmetric, each rebinding both
'     cells' lists rather than assuming only one drives the other. See
'     docs/ARCHITECTURE.md §7.2 for the design and a note on the one real
'     bug this surfaced (Range.Value on a multi-cell Union is an array, not
'     a scalar - AutoFillIfSingle has to walk target.Cells, not read
'     target.Value directly).
'
' 0.8.1 - phase 8's original scope, finally built (see the 0.8.0 note below,
' which explicitly deferred this):
'
'   - modSettings.CurrencySymbol (already reading SET_CURRENCY, previously
'     unreferenced) is now read everywhere a NumberFormat or Format$ used to
'     hardcode ChrW(163) - modReports (Summary and Reports), modExport's
'     export-report snapshot, and modJobs.RemoveRow's delete-confirmation
'     dialog. The NumberFormat sites go through the new
'     modSettings.CurrencyFormatCode rather than CurrencySymbol() directly -
'     verified directly against this workbook over COM: setting a Range's
'     NumberFormat from VBA to a BARE or quoted currency symbol (e.g.
'     "$#,##0.00") is silently canonicalised by Excel back to the OS's own
'     regional currency symbol, no matter what SET_CURRENCY holds - invisible
'     on a GBP-locale machine, where it happened to always match, but it
'     would have silently defeated the setting everywhere else. Excel's
'     "[$symbol]" bracket syntax round-trips exactly as given and was
'     confirmed to render identically for the unaffected GBP case, so
'     CurrencyFormatCode is a strict fix, not a behaviour change. Format$
'     (RemoveRow's dialog) does not have this problem - confirmed separately
'     - and keeps using CurrencySymbol() plain. Changing SET_CURRENCY still
'     only repaints Summary/Reports on the next InitialiseWorkbook run, same
'     as every other piece of their formatting.
'   - Conditional formatting exists in the workbook for the first time.
'     modInit.ApplyStatusFormat gives every location sheet's Status column an
'     amber fill whenever a row reads anything other than "OK" - the
'     "warning state" §5/§11 already called out as "conditionally
'     formatted". modReports.FormatSummaryErrors gives Summary's Type/Family
'     columns red bold text when a job references a paper stock that no
'     longer exists in tblPapers ("(not in Papers)") - the "error state",
'     matched to the one place the workbook's own formulas already flag a
'     genuine data-integrity break rather than a routine per-row warning.
'   - modReports.DrawLegend adds the Summary-sheet legend §11 planned and
'     never built: the four everyday cell-role colours (input, calculated,
'     configuration, read-only reference), at O9 down, in the colours
'     already used for those roles elsewhere on Summary/Reports.
'
' 0.8.0 - NextId high-water mark, Export/Import, Reports rework, UI snag
' list. Real new functionality, not just fixes, so the phase digit moves:
'
'   - modRegistry.NextJobId now allocates from a persisted high-water mark
'     (a new "Job ID HWM" registry column) instead of a scan of rows on the
'     sheet, so deleting the highest-numbered job can no longer reissue its
'     ID. Fixed first because Export/Import both depend on the "globally
'     unique, never re-issued" invariant actually holding.
'   - modExport gained Export All Locations (every room in one pass) and an
'     Export report snapshot - a static-value .xlsx of the Reports sheet's
'     current filtered/sorted view, with its own signature stamped on
'     Reports for the delete command below to check against. Exports now
'     resolve to a selectable folder (SET_EXPORT_FOLDER, falling back to the
'     workbook's own location if unset or not present on this machine), and
'     the OneDrive path fix now also searches
'     ~/Library/CloudStorage/OneDrive* on Mac, where the Windows-only
'     OneDrive environment variables never existed in the first place.
'   - modImport is new: restores or merges an exported file into a
'     location's job table. Rows match by Job ID - append if new, overwrite
'     if not - and only input/snapshot columns are written, so an imported
'     row costs correctly from its own frozen rates regardless of the
'     target workbook's current catalogue.
'   - Cost Calculations renamed to Reports: Technician/Printer/Paper
'     Stock/Quantity filters, sort-by-any-column, a hidden Job ID
'     correlation column, and a bulk "Delete visible records" command
'     confined to whatever the active filters show. Summary regrouped to a
'     Location x Printer x Paper stock key.
'   - The 2026-09-21 snag list: no default sheet-protection password,
'     unlocked configuration table inputs, freeze panes moved off the
'     config sheets and onto Reports, tab order fixed, a hide/show toggle
'     for the config sheets, column and row groups, Settings notes
'     wrapped, and catalogue row add/remove buttons.
'
' Not part of this bump: phase 8's original scope (the Summary legend,
' conditional formatting for warning/error states, and wiring SET_CURRENCY
' into the still-hardcoded GBP symbol) remains outstanding.
'
' 0.7.1 - review pass before phase 8. Seven fixes, no new features:
'
'   - modReports.SheetNamed moved a sheet LATER with Move Before:, which is a
'     no-op, so every build shipped Cost Calculations on tab 1 and Summary on
'     tab 2 - the reverse of what the code asks for.
'   - Val() on numeric cells replaced with modUtils.NumOf. Val round-trips
'     through a locale-formatted string, so on any comma-decimal machine every
'     paper cost, consumable rate and stock dimension read as zero. Design 8.2
'     states this rule for dates; the cause is not specific to them.
'   - modSnapshot.LogAudit and modExport.ExportSig no longer wrap a batch of
'     independent operations in one On Error Resume Next (design 8.2).
'   - modRegistry.WriteRegistry no longer raises on a workbook with no print
'     rooms, which is the design 3.3 case it most needs to survive.
'   - modSettings.Money made locale-safe and given CurrencySymbol.
'
' A patch increment, not a phase: the phase digit still reads 7 because phase
' 8 has not been built. The revision digit is what "0.<phase>.<revision>"
' exists for.
Public Const APP_VERSION As String = "0.9.0"
Public Const APP_AUTHOR As String = "Adam Shailer"

Public Function VersionString() As String
    VersionString = APP_NAME & " v" & APP_VERSION & " (data schema " & SCHEMA_VER & ")"
End Function

' --------------------------------------------------------------- stamping ---
' Called by build.ps1 AFTER InitialiseWorkbook, not by it.
'
' "Built" has to mean when this file was produced, and InitialiseWorkbook is
' re-run by users every time they add a print room. Stamping from there would
' quietly reset the build date to whenever somebody last duplicated a sheet.
Public Sub StampBuild()
    On Error GoTo Fail
    AppOff
    EnsureVersionSettings
    SetSetting "BUILT", Now
    SetSetting "BUILT_BY", CurrentUser
    WriteAbout
    StampProperties
    AppOn
    Say "Stamped as " & VersionString & ".", "Built " & Format$(Now, "dd/mm/yyyy hh:mm") & " by " & CurrentUser & "."
    Exit Sub
Fail:
    AppReset
    ReportError "StampBuild"
End Sub

' Adds the version rows to tblSettings if they are not there, and gives each a
' workbook-scoped SET_ name so modSettings can reach it like any other setting.
'
' APP_VER is written from the constant every time, because the code in the
' workbook IS its version - importing a newer module set makes it a newer
' workbook, and the cell should say so without waiting for a rebuild. BUILT
' and BUILT_BY are left to StampBuild.
Public Sub EnsureVersionSettings()
    EnsureSetting "APP_VER", "Workbook version", "Read-only. Matches the code in this workbook."
    EnsureSetting "BUILT", "Built", "Read-only. When this file was produced."
    EnsureSetting "BUILT_BY", "Built by", "Read-only."
    SetSetting "APP_VER", APP_VERSION
End Sub

' Public: modExport.EnsureExportSettings reuses this for the same reason
' EnsureVersionSettings' own rows go through it - a setting is not real to
' modSettings.SettingText until it has a SET_<KEY> name, and this is the one
' place that adds a row to tblSettings and names it in the same step.
Public Function EnsureSetting(ByVal Key As String, ByVal Label As String, ByVal Notes As String) As Range
    Dim lo As ListObject, i As Long, r As ListRow, c As Range

    Set lo = Tbl("tblSettings")
    If lo Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Key").Value)), Key, vbTextCompare) = 0 Then
            Set c = CellIn(lo, i, "Value")
            Exit For
        End If
    Next i

    If c Is Nothing Then
        UnlockSheet lo.Parent
        Set r = lo.ListRows.Add
        r.Range.Cells(1, 1).Value = Key
        r.Range.Cells(1, 2).Value = Label
        r.Range.Cells(1, 4).Value = Notes
        Set c = r.Range.Cells(1, 3)
        RelockSheet lo.Parent
    End If

    ' modSettings addresses every setting as SET_<KEY>, so the name is what
    ' makes a new row a real setting rather than just a row of text.
    On Error Resume Next
    ThisWorkbook.Names("SET_" & Key).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="SET_" & Key, _
        RefersTo:="='" & c.Parent.Name & "'!" & c.Address(True, True, xlA1)

    Set EnsureSetting = c
End Function

' ------------------------------------------------------------ about block ---
' Sits below whatever the Settings sheet's tables occupy, so adding settings
' later pushes it down rather than writing over it. Rewritten in full each
' time, which is why it clears its own area first.
Public Sub WriteAbout()
    Dim ws As Worksheet, r As Long, lo As ListObject, bottom As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    For Each lo In ws.ListObjects
        If lo.Range.Row + lo.Range.Rows.Count - 1 > bottom Then bottom = lo.Range.Row + lo.Range.Rows.Count - 1
    Next lo
    r = bottom + 3

    UnlockSheet ws
    ws.Range(ws.Cells(bottom + 1, 1), ws.Cells(bottom + 20, 4)).Clear

    ws.Cells(r, 1).Value = "About"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12

    Lbl ws, r + 1, "Workbook", APP_NAME
    ' Read from the settings rows rather than the constants, so the block can
    ' never disagree with what the rest of the workbook reports.
    Lbl ws, r + 2, "Version", "=SET_APP_VER"
    Lbl ws, r + 3, "Data schema", "=SET_SCHEMA"
    Lbl ws, r + 4, "Built", "=TEXT(SET_BUILT,""dd/mm/yyyy hh:mm"")"
    Lbl ws, r + 5, "Built by", "=SET_BUILT_BY"

    ws.Cells(r + 7, 1).Value = "Logs print output and cost per student across the university's print bureaux. " & _
        "Each record freezes the prices it was costed at, so changing a paper or consumable price " & _
        "never alters what has already been charged."
    ws.Cells(r + 8, 1).Value = "Version numbering is 0.<build phase>.<revision>, tracking the build sequence " & _
        "in the Stage 1 design document. The data schema version is separate and changes only when " & _
        "the print job table's columns do."
    ws.Cells(r + 9, 1).Value = "Specified by " & APP_AUTHOR & ". Built by Claude (Anthropic) working with " & _
        APP_AUTHOR & ", September 2026."

    ' Left unmerged and unwrapped so the text simply overflows to the right.
    ' Merged cells do not AutoFit, so wrapping here would need a hard-coded row
    ' height that breaks at a different zoom or font.
    Dim i As Long
    For i = 7 To 9
        ws.Cells(r + i, 1).Font.Italic = True
    Next i

    RelockSheet ws
End Sub

Private Sub Lbl(ByVal ws As Worksheet, ByVal RowNo As Long, ByVal Label As String, ByVal Value As String)
    ws.Cells(RowNo, 1).Value = Label
    ws.Cells(RowNo, 1).Font.Bold = True
    If Left$(Value, 1) = "=" Then
        ws.Cells(RowNo, 3).Formula = Value
    Else
        ws.Cells(RowNo, 3).Value = Value
    End If
End Sub

' --------------------------------------------------- file-level identity ---
' So the version is visible without opening the file: Explorer's details pane,
' Finder's Get Info, and File > Info in Excel all read these.
' Each property is set individually and failures are reported rather than
' swallowed. An earlier version wrapped the lot in On Error Resume Next and
' silently wrote nothing at all - the properties came back empty and the only
' way to find out was to go and look.
Private Sub StampProperties()
    Dim bad As String
    bad = SetProp("Title", APP_NAME & " v" & APP_VERSION)
    bad = bad & SetProp("Subject", "Print output and cost logging, data schema " & SCHEMA_VER)
    bad = bad & SetProp("Author", APP_AUTHOR)
    bad = bad & SetProp("Comments", VersionString & ", built " & Format$(Now, "dd/mm/yyyy hh:mm"))
    If Len(bad) > 0 Then
        Say "Some file properties could not be set.", bad, _
            "The version is still on the Settings sheet; only the Windows file details are affected."
    End If
End Sub

Private Function SetProp(ByVal Nm As String, ByVal Value As String) As String
    On Error GoTo Fail
    ThisWorkbook.BuiltinDocumentProperties(Nm).Value = Value
    Exit Function
Fail:
    SetProp = "- " & Nm & ": " & Err.Number & " " & Err.Description & vbCrLf
End Function

' ------------------------------------------------------------------ about ---
' Bound to the About button on the Settings sheet.
Public Sub ShowAbout()
    Dim built As String, who As String, ver As String

    ver = SettingText("APP_VER", "(not stamped)")
    built = SettingText("BUILT")
    who = SettingText("BUILT_BY")

    Say APP_NAME & " v" & ver, _
        "Data schema " & SCHEMA_VER & IIf(Len(built) > 0, vbCrLf & "Built " & built, "") & _
        IIf(Len(who) > 0, " by " & who, "") & vbCrLf & vbCrLf & _
        "Specified by " & APP_AUTHOR & ". Built by Claude (Anthropic) working with " & APP_AUTHOR & ", September 2026.", _
        "Full details are in the About section at the bottom of this sheet.", _
        "About"
End Sub
