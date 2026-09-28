Attribute VB_Name = "modCatalog"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Configuration loaded once per operation.
'
' Reports and validation sweeps touch every job row; re-reading the Papers and
' Printers tables inside those loops is the classic way to make a workbook feel
' broken. Everything is read once into clsDict instances and looked up here.
'
' Stocks and printers are objects (clsStock, clsPrinterDef) rather than Public
' Types. VBA will not put a user-defined type declared in a standard module
' into a Variant or a Collection, and clsDict is a Collection underneath.
'
' A lookup that finds nothing returns an object with Found = False rather than
' Nothing, so callers can test .Found without guarding every reference.

Private mStocks As clsDict      ' by description
Private mPrinters As clsDict    ' by model
Private mTechs As clsDict       ' by name -> TechID
Private mBasis As clsDict       ' family -> "Sheet" / "Roll"
Private mLoaded As Boolean
Private mDirty As Boolean       ' set alongside mLoaded=False; cleared once
                                 ' every print room's dropdowns have actually
                                 ' been rebound (ThisWorkbook.Workbook_
                                 ' SheetDeactivate) - see modLists.
                                 ' RebindAllLocationDropdowns. Invalidate on
                                 ' its own only clears the in-memory cache;
                                 ' nothing pushes the change out to the
                                 ' dropdowns already drawn on each sheet until
                                 ' something rebinds them, and this flag is
                                 ' how that "something" knows it is needed.

Public Sub Invalidate()
    mLoaded = False
    mDirty = True
End Sub

' True from the moment a catalogue edit invalidates the cache until
' RebindAllLocationDropdowns next runs - i.e. "some print room's Printer/
' Paper Stock/Technician dropdown may now be showing stale data".
Public Function CatalogDirty() As Boolean
    CatalogDirty = mDirty
End Function

Public Sub ClearCatalogDirty()
    mDirty = False
End Sub

Public Sub LoadCatalog(Optional ByVal Force As Boolean = False)
    If mLoaded And Not Force Then Exit Sub

    Dim lo As ListObject, i As Long
    Dim s As clsStock, p As clsPrinterDef
    Dim szW As Double, szH As Double

    Set mStocks = New clsDict
    Set mPrinters = New clsDict
    Set mTechs = New clsDict
    Set mBasis = New clsDict

    Set lo = Tbl("tblPaperFamilies")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Family").Value) > 0 Then
            mBasis.Add CStr(CellIn(lo, i, "Family").Value), CStr(CellIn(lo, i, "Measurement basis").Value)
        End If
    Next i

    Set lo = Tbl("tblPapers")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Description").Value) > 0 Then
            Set s = New clsStock
            s.StockID = CStr(CellIn(lo, i, "StockID").Value)
            s.Description = CStr(CellIn(lo, i, "Description").Value)
            s.PaperType = CStr(CellIn(lo, i, "Paper type").Value)
            s.Family = CStr(CellIn(lo, i, "Family").Value)
            s.Measure = CStr(mBasis.Item(s.Family))
            s.WidthMM = NumOf(CellIn(lo, i, "Width mm"))
            s.HeightMM = NumOf(CellIn(lo, i, "Height mm"))
            s.Cost = NumOf(CellIn(lo, i, "Cost"))
            s.Active = (StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0)
            If ColumnExists(lo, "Supplied by student") Then
                s.SuppliedByStudent = (StrComp(CStr(CellIn(lo, i, "Supplied by student").Value), "Yes", vbTextCompare) = 0)
            End If
            s.Found = True
            mStocks.Add s.Description, s
        End If
    Next i

    Set lo = Tbl("tblPrinters")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Model").Value) > 0 Then
            Set p = New clsPrinterDef
            p.PrinterID = CStr(CellIn(lo, i, "PrinterID").Value)
            p.Model = CStr(CellIn(lo, i, "Model").Value)
            p.Consumable = CStr(CellIn(lo, i, "Consumable type").Value)
            p.RatePerM2 = NumOf(CellIn(lo, i, "Cost per m2"))
            If ColumnExists(lo, "Max roll width mm") Then
                p.MaxRollWidthMM = NumOf(CellIn(lo, i, "Max roll width mm"))
            End If
            If ColumnExists(lo, "Max sheet size") Then
                p.MaxSheetSize = Trim$(CStr(CellIn(lo, i, "Max sheet size").Value))
                If Len(p.MaxSheetSize) > 0 Then
                    ' Via local scratch variables, not p.MaxSheetWidthMM/
                    ' HeightMM directly - an object's public field passed as
                    ' a ByRef argument doesn't reliably write back over COM,
                    ' confirmed directly: StdSizeDims returned True with the
                    ' correct values, but p.MaxSheetWidthMM/HeightMM stayed 0
                    ' every time until routed through plain local Doubles.
                    If StdSizeDims(p.MaxSheetSize, szW, szH) Then
                        p.MaxSheetWidthMM = szW
                        p.MaxSheetHeightMM = szH
                    End If
                End If
            End If
            p.Active = (StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0)
            p.Found = True
            mPrinters.Add p.Model, p
        End If
    Next i

    Set lo = Tbl("tblTechnicians")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Name").Value) > 0 Then
            mTechs.Add CStr(CellIn(lo, i, "Name").Value), CStr(CellIn(lo, i, "TechID").Value)
        End If
    Next i

    mLoaded = True
End Sub

Public Function Stock(ByVal Description As String) As clsStock
    LoadCatalog
    If mStocks.Exists(Description) Then
        Set Stock = mStocks.Obj(Description)
    Else
        Set Stock = New clsStock
    End If
End Function

Public Function Prn(ByVal Model As String) As clsPrinterDef
    LoadCatalog
    If mPrinters.Exists(Model) Then
        Set Prn = mPrinters.Obj(Model)
    Else
        Set Prn = New clsPrinterDef
    End If
End Function

Public Function TechID(ByVal TechName As String) As String
    LoadCatalog
    TechID = CStr(mTechs.Item(TechName))
End Function

Public Function Basis(ByVal Family As String) As String
    LoadCatalog
    Basis = CStr(mBasis.Item(Family))
End Function

' A stock is usable on a printer only when the printer can take that kind of
' stock (roll/sheet, via Max roll width mm / Max sheet size) AND the stock's
' own size fits within that capacity. Family no longer plays any part here -
' it used to stand in for a width band (families called "Short Roll"/"Long
' Roll"), which is exactly the coarseness this rework replaces with a real
' numeric fit check (printer/paper compatibility rework plan).
'
' The two "Supplied (Roll)"/"Supplied (Sheet)" catalogue rows (student-
' supplied stock) have no fixed size to check here - the real dimensions
' aren't known until job entry (Print Width mm / the job-row Sheet size
' column), so compatibility for them is just "the printer takes this kind of
' stock at all". modValidation re-checks the row's own entered size against
' the chosen printer's capacity once it exists.
Public Function Compatible(ByVal Model As String, ByVal Description As String) As Boolean
    Dim p As clsPrinterDef, s As clsStock
    Set p = Prn(Model)
    Set s = Stock(Description)
    If Not p.Found Then Exit Function
    If Not s.Found Then Exit Function

    If s.Measure = "Roll" Then
        If p.MaxRollWidthMM <= 0 Then Exit Function
        If s.SuppliedByStudent Then
            Compatible = True
        Else
            Compatible = (s.WidthMM <= p.MaxRollWidthMM)
        End If
    ElseIf s.Measure = "Sheet" Then
        If Len(p.MaxSheetSize) = 0 Then Exit Function
        If s.SuppliedByStudent Then
            Compatible = True
        Else
            Compatible = FitsWithinMaxSheet(s.WidthMM, s.HeightMM, p.MaxSheetWidthMM, p.MaxSheetHeightMM)
        End If
    End If
End Function

' Whether a W x H stock fits within a MaxW x MaxH capacity in EITHER
' orientation - a smaller sheet can be fed either way round, so the
' comparison normalises both pairs to (longer side, shorter side) rather
' than assuming a fixed portrait/landscape match. Handles non-standard
' shapes (a Custom sheet stock with no Std. size name) exactly the same way
' as a named standard size, since it only ever looks at Width mm/Height mm.
Public Function FitsWithinMaxSheet(ByVal W As Double, ByVal H As Double, ByVal MaxW As Double, ByVal MaxH As Double) As Boolean
    If MaxW <= 0 Or MaxH <= 0 Then Exit Function
    If W <= 0 Or H <= 0 Then Exit Function
    FitsWithinMaxSheet = (Application.WorksheetFunction.Max(W, H) <= Application.WorksheetFunction.Max(MaxW, MaxH)) _
                      And (Application.WorksheetFunction.Min(W, H) <= Application.WorksheetFunction.Min(MaxW, MaxH))
End Function

' A human-readable description of what a printer can take, for the AT-05
' "incompatible combination" messages in modValidation - replaces the old
' "does not support the X paper family" phrasing now that compatibility is
' a size fit rather than family membership.
Public Function CapacityText(ByVal p As clsPrinterDef) As String
    Dim parts As String
    If p.MaxRollWidthMM > 0 Then parts = "roll stock up to " & Format$(p.MaxRollWidthMM, "#,##0") & " mm wide"
    If Len(p.MaxSheetSize) > 0 Then
        If Len(parts) > 0 Then parts = parts & " or "
        parts = parts & "sheet stock up to " & p.MaxSheetSize
    End If
    If Len(parts) = 0 Then parts = "no paper stock at all"
    CapacityText = parts
End Function

' Looks up a Standard Sizes row by name - shared by printer capacity
' resolution (LoadCatalog, above) and a student-supplied sheet job's own
' "Sheet size" cell (modSnapshot.StampRow, modValidation.OnSheetSizeChanged).
' Returns False (leaving W/H untouched) when the name isn't found, e.g. a
' size removed from the catalogue after being chosen.
Public Function StdSizeDims(ByVal SizeName As String, ByRef W As Double, ByRef H As Double) As Boolean
    Dim lo As ListObject, i As Long
    Set lo = Tbl("tblStandardSizes")
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Size name").Value)), Trim$(SizeName), vbTextCompare) = 0 Then
            W = NumOf(CellIn(lo, i, "Width mm"))
            H = NumOf(CellIn(lo, i, "Height mm"))
            StdSizeDims = True
            Exit Function
        End If
    Next i
End Function

' Active stocks compatible with the printer, for the dependent dropdown.
Public Function StocksFor(ByVal Model As String) As Collection
    Dim out As Collection, k As Variant, s As clsStock
    Set out = New Collection
    LoadCatalog
    For Each k In mStocks.Keys
        Set s = mStocks.Obj(CStr(k))
        If s.Active Then
            If Compatible(Model, s.Description) Then out.Add s.Description
        End If
    Next k
    Set StocksFor = out
End Function

' Active printers permitted at this location (spec 9.1, AT-06).
Public Function PrintersFor(ByVal ws As Worksheet) As Collection
    Dim out As Collection, k As Variant, p As clsPrinterDef, allowed As String
    Set out = New Collection
    LoadCatalog
    allowed = LocValue(ws, "LOC_Printers")
    For Each k In mPrinters.Keys
        Set p = mPrinters.Obj(CStr(k))
        If p.Active Then
            If InList(allowed, p.Model) Then out.Add p.Model
        End If
    Next k
    Set PrintersFor = out
End Function

' The reverse of PrintersFor/StocksFor: active printers permitted at this
' location whose supported families include the given stock's family (spec
' 1a). Paper Stock can now be chosen before Printer, so Printer's own list
' has to be narrowable by the row's stock just as Stock's list is narrowable
' by the row's printer.
Public Function PrintersForStock(ByVal ws As Worksheet, ByVal Description As String) As Collection
    Dim out As Collection, k As Variant, p As clsPrinterDef, allowed As String
    Set out = New Collection
    LoadCatalog
    allowed = LocValue(ws, "LOC_Printers")
    For Each k In mPrinters.Keys
        Set p = mPrinters.Obj(CStr(k))
        If p.Active Then
            If InList(allowed, p.Model) Then
                If Compatible(p.Model, Description) Then out.Add p.Model
            End If
        End If
    Next k
    Set PrintersForStock = out
End Function

Public Function ActiveTechnicians() As Collection
    Dim out As Collection, lo As ListObject, i As Long
    Set out = New Collection
    Set lo = Tbl("tblTechnicians")
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0 Then
            If Len(CellIn(lo, i, "Name").Value) > 0 Then
                out.Add CStr(CellIn(lo, i, "Name").Value)
            End If
        End If
    Next i
    Set ActiveTechnicians = out
End Function

' Every active printer/paper stock DEFINED in the catalogue, regardless of
' whether it has been used in a job yet - for the Reports page filters (snag
' list item 4), which deliberately do not restrict one by the other the way
' StocksFor does for job entry. Sorted, unlike PrintersFor/StocksFor, because
' nothing else about those two orders it for a caller and a filter dropdown
' is read by a person.
Public Function AllActivePrinters() As Collection
    Dim out As New Collection, k As Variant, p As clsPrinterDef
    LoadCatalog
    For Each k In mPrinters.Keys
        Set p = mPrinters.Obj(CStr(k))
        If p.Active Then out.Add p.Model
    Next k
    Set AllActivePrinters = SortedTextCollection(out)
End Function

Public Function AllActiveStocks() As Collection
    Dim out As New Collection, k As Variant, s As clsStock
    LoadCatalog
    For Each k In mStocks.Keys
        Set s = mStocks.Obj(CStr(k))
        If s.Active Then out.Add s.Description
    Next k
    Set AllActiveStocks = SortedTextCollection(out)
End Function

' ---------------------------------------------------- catalogue row edit ---
' Add/remove a row on one of the catalogue tables - Printers, Papers, Print
' Technicians (snag list item 17). Mirrors modJobs.AddPrintJob/RemoveRow:
' ListRows.Add keeps the new row's formatting and validation, because an
' Excel Table extends both automatically, and removing asks first and shows
' what is about to go - the same "never a bare are you sure" rule spec 10.12
' sets for job rows.
'
' Deleting a catalogue row never touches a job already recorded against it:
' every job snapshots the price it was costed at (the S_* columns), which is
' the whole point of the snapshot - see modExport's "content is the
' irreversible decision" note.
Public Sub AddCatalogRow(ByVal TableName As String)
    Dim lo As ListObject, r As ListRow, ws As Worksheet
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    AppOff
    UnlockSheet ws
    If lo.ListRows.Count = 1 And IsBlankRow(lo, 1) Then
        Set r = lo.ListRows(1)
    Else
        Set r = lo.ListRows.Add
    End If
    RelockSheet ws
    Invalidate
    AppOn

    r.Range.Cells(1, 1).Select
End Sub

Public Sub RemoveCatalogRow(ByVal TableName As String)
    Dim lo As ListObject, ws As Worksheet, n As Long, c As Range
    Dim detail As String, i As Long

    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    If lo.DataBodyRange Is Nothing Then
        Say "There is nothing to remove.", "'" & TableName & "' has no rows."
        Exit Sub
    End If

    Set c = Application.Intersect(Selection.Cells(1, 1).EntireRow, lo.DataBodyRange)
    If c Is Nothing Then
        Say "No row is selected.", "This command works on the row the cursor is in.", "Click any cell in the row you want to remove, then try again."
        Exit Sub
    End If
    n = c.Row - lo.DataBodyRange.Row + 1

    For i = 1 To lo.ListColumns.Count
        If i > 1 Then detail = detail & "  "
        detail = detail & lo.ListColumns(i).Name & ": " & CStr(lo.DataBodyRange.Cells(n, i).Value)
    Next i

    If Not Ask("Remove this row?" & vbCrLf & vbCrLf & detail & vbCrLf & vbCrLf & _
        "Print jobs already recorded against it keep their frozen prices and are not affected. This cannot be undone.", _
        "Remove row") Then Exit Sub

    AppOff
    UnlockSheet ws
    lo.ListRows(n).Delete
    RelockSheet ws
    Invalidate
    AppOn
End Sub

' A pure rename (0.9.14, "shorten headers to save space"), same reasoning as
' modInit.EnsureQtyColumnName: tblPapers ships in PrintCosts.xlsx with this
' column still called "Standard size" (§16.1 - schema-shape changes are
' self-provisioned at setup rather than hand-edited into the static .xlsx),
' and nothing in this codebase looks it up by name (unlike Quantity, this one
' needed no other call sites touched), so the rename is free of any lookup
' risk. Checked first, like every other Ensure* in this project.
Public Sub EnsureStdSizeColumnName()
    Dim lo As ListObject, i As Long
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub

    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, "Std. size", vbTextCompare) = 0 Then Exit Sub
    Next i

    UnlockSheet lo.Parent
    On Error Resume Next
    lo.ListColumns("Standard size").Name = "Std. size"
    On Error GoTo 0
    RelockSheet lo.Parent
End Sub

' ---------------------------------------------- printer capacity columns ---
' Adds the two capacity fields Compatible() now uses in place of the old
' "Supported families" multi-select: a printer accepts roll stock iff Max
' roll width mm is set, sheet stock iff Max sheet size is set - both can be
' set on the same printer. Checked-first/idempotent like every other Ensure*
' in this project. Called once from InitialiseWorkbook, not per-location -
' this is a catalogue-level table, not a job table.
' A workbook-scoped name for tblStandardSizes[Size name], re-created (delete
' then re-add, same idiom modInit.EnsureLocName uses) rather than referenced
' directly as a structured reference in a cross-sheet Data Validation list:
' Validation.Add's Formula1 rejects a bare cross-sheet structured reference
' with a plain 1004 over COM automation, even though the same text works
' fine typed into the dialog by hand - a defined name is what makes a
' cross-sheet list source reliable here, same reasoning modLists.bas gives
' for why every OTHER dropdown in this project stages its list first rather
' than pointing straight at a table range.
Public Sub EnsureStdSizesName()
    On Error Resume Next
    ThisWorkbook.Names("RNG_STD_SIZES").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="RNG_STD_SIZES", RefersTo:="=tblStandardSizes[Size name]"
End Sub

Public Sub EnsurePrinterCapacityColumns()
    Dim lo As ListObject, lc As ListColumn, ws As Worksheet
    Set lo = Tbl("tblPrinters")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    EnsureStdSizesName

    If Not ColumnExists(lo, "Max roll width mm") Then
        UnlockSheet ws
        Set lc = lo.ListColumns.Add(ColIdx(lo, "Cost per m2") + 1)
        lc.Name = "Max roll width mm"
        If Not lc.DataBodyRange Is Nothing Then
            StyleInputCell lc.DataBodyRange
            With lc.DataBodyRange.Validation
                .Delete
                .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
                .IgnoreBlank = True
                .InputTitle = "Max roll width mm"
                .InputMessage = "The widest roll stock this printer can take, in millimetres. Leave blank if this printer cannot print on roll stock."
                .ErrorTitle = "Max roll width mm"
                .ErrorMessage = "Enter 0 or more, in millimetres."
            End With
        End If
        RelockSheet ws
    End If

    Set lo = Tbl("tblPrinters")
    If Not ColumnExists(lo, "Max sheet size") Then
        UnlockSheet ws
        Set lc = lo.ListColumns.Add(ColIdx(lo, "Max roll width mm") + 1)
        lc.Name = "Max sheet size"
        If Not lc.DataBodyRange Is Nothing Then
            StyleInputCell lc.DataBodyRange
            With lc.DataBodyRange.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=RNG_STD_SIZES"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = "Max sheet size"
                .InputMessage = "The largest standard sheet size this printer can take. Leave blank if this printer cannot print on sheet stock."
                .ErrorTitle = "Max sheet size"
                .ErrorMessage = "Choose one of the standard sizes listed on Settings."
            End With
        End If
        RelockSheet ws
    End If

    EnsurePrinterCapacityNote
End Sub

' The Printers-sheet note (H4) explained the old Select families... button;
' this asserts fresh wording every run rather than leaving PrintCosts.xlsx's
' shipped text to go stale now that button is gone (modPicker.PickFamilies,
' removed) - the same "self-heal via VBA, don't hand-edit the source file"
' reasoning every other note/label in this project follows.
Private Sub EnsurePrinterCapacityNote()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Printers")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    UnlockSheet ws
    ws.Range("H4").Value = "Max roll width mm / Max sheet size set what each printer can take - blank means it cannot print that kind of stock at all. A paper stock can only be used on a printer whose capacity it fits within."
    RelockSheet ws
End Sub

' One-time, idempotent: derives each printer's new Max roll width mm / Max
' sheet size from whatever it was compatible with under the OLD family-list
' rule, so migrating an existing workbook doesn't silently change which
' stocks a printer can already use - a principled translation of the old
' semantics into the new model, not a guess. Must run BEFORE
' MigrateRollFamilies (below), while tblPapers' Family values still say
' "Short Roll"/"Long Roll" and can still be matched against the legacy
' Supported families text; also removes that column once every printer has
' been read, which is what makes this whole Sub a true one-off - the next
' run exits immediately on the missing column (ColumnExists check, below).
Public Sub MigratePrinterCapacities()
    Dim lo As ListObject, i As Long, legacy As String
    Dim maxRoll As Double, sizeName As String

    Set lo = Tbl("tblPrinters")
    If lo Is Nothing Then Exit Sub
    If Not ColumnExists(lo, "Supported families") Then Exit Sub
    If Not ColumnExists(lo, "Max roll width mm") Then Exit Sub
    If Not ColumnExists(lo, "Max sheet size") Then Exit Sub

    UnlockSheet lo.Parent
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Model").Value) > 0 Then
            legacy = CStr(CellIn(lo, i, "Supported families").Value)
            If Len(legacy) > 0 Then
                If NumOf(CellIn(lo, i, "Max roll width mm")) = 0 And Len(Trim$(CStr(CellIn(lo, i, "Max sheet size").Value))) = 0 Then
                    maxRoll = WidestLegacyRollStock(legacy)
                    If maxRoll > 0 Then CellIn(lo, i, "Max roll width mm").Value = maxRoll

                    sizeName = SmallestLegacySheetSize(legacy)
                    If Len(sizeName) > 0 Then CellIn(lo, i, "Max sheet size").Value = sizeName
                End If
            End If
        End If
    Next i

    lo.ListColumns("Supported families").Delete
    RelockSheet lo.Parent
    Invalidate
End Sub

' The widest Roll-family stock the old family-list rule let this printer
' take - becomes its Max roll width mm, so nothing it could already print
' stops fitting once the check switches from family membership to a real
' size comparison.
Private Function WidestLegacyRollStock(ByVal LegacyFamilies As String) As Double
    Dim lo As ListObject, i As Long, fam As String, w As Double
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Description").Value) > 0 Then
            fam = CStr(CellIn(lo, i, "Family").Value)
            If InList(LegacyFamilies, fam) Then
                If StrComp(Basis(fam), "Roll", vbTextCompare) = 0 Then
                    w = NumOf(CellIn(lo, i, "Width mm"))
                    If w > WidestLegacyRollStock Then WidestLegacyRollStock = w
                End If
            End If
        End If
    Next i
End Function

' The smallest catalogue Standard Size whose bounding box (either
' orientation) contains every Sheet-family stock the old family-list rule
' let this printer take - becomes its Max sheet size, same "preserve
' current behaviour" principle as WidestLegacyRollStock above. Every
' compatible stock's own (longer side, shorter side) is tracked
' independently and the smallest standard size covering the worst case on
' each axis is chosen, so a printer that could already take a tall narrow
' sheet and a short wide one both still fit afterward.
Private Function SmallestLegacySheetSize(ByVal LegacyFamilies As String) As String
    Dim lo As ListObject, i As Long, fam As String, w As Double, h As Double
    Dim needLong As Double, needShort As Double

    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Description").Value) > 0 Then
            fam = CStr(CellIn(lo, i, "Family").Value)
            If InList(LegacyFamilies, fam) Then
                If StrComp(Basis(fam), "Sheet", vbTextCompare) = 0 Then
                    w = NumOf(CellIn(lo, i, "Width mm"))
                    h = NumOf(CellIn(lo, i, "Height mm"))
                    If Application.WorksheetFunction.Max(w, h) > needLong Then needLong = Application.WorksheetFunction.Max(w, h)
                    If Application.WorksheetFunction.Min(w, h) > needShort Then needShort = Application.WorksheetFunction.Min(w, h)
                End If
            End If
        End If
    Next i
    If needLong = 0 Then Exit Function

    Dim sl As ListObject, j As Long, sw As Double, shh As Double, slng As Double, sshrt As Double
    Dim bestArea As Double, thisArea As Double, bestName As String
    Set sl = Tbl("tblStandardSizes")
    If sl Is Nothing Then Exit Function
    For j = 1 To sl.ListRows.Count
        If Len(CellIn(sl, j, "Size name").Value) > 0 Then
            sw = NumOf(CellIn(sl, j, "Width mm"))
            shh = NumOf(CellIn(sl, j, "Height mm"))
            slng = Application.WorksheetFunction.Max(sw, shh)
            sshrt = Application.WorksheetFunction.Min(sw, shh)
            If slng >= needLong And sshrt >= needShort Then
                thisArea = slng * sshrt
                If Len(bestName) = 0 Or thisArea < bestArea Then
                    bestArea = thisArea
                    bestName = CStr(CellIn(sl, j, "Size name").Value)
                End If
            End If
        End If
    Next j
    SmallestLegacySheetSize = bestName
End Function

' ------------------------------------------------------- paper formats -----
' Merges the two legacy width-band families ("Short Roll", "Long Roll") into
' a single "Roll" family, now that Compatible() no longer needs one family
' per width band - families only need to distinguish Sheet from Roll stock
' (printer/paper compatibility rework plan). Idempotent: exits immediately
' once neither legacy name is present, so this is a true one-off per
' workbook. Must run before EnsureSuppliedByStudentColumn (below), which
' seeds a stock row against the "Roll" family this creates.
Public Sub MigrateRollFamilies()
    Dim fl As ListObject, pl As ListObject, i As Long, fam As String
    Dim needsMigration As Boolean, keepRow As Long, removeRow As Long

    Set fl = Tbl("tblPaperFamilies")
    If fl Is Nothing Then Exit Sub

    For i = 1 To fl.ListRows.Count
        fam = Trim$(CStr(CellIn(fl, i, "Family").Value))
        If StrComp(fam, "Short Roll", vbTextCompare) = 0 Or StrComp(fam, "Long Roll", vbTextCompare) = 0 Then
            needsMigration = True
            Exit For
        End If
    Next i
    If Not needsMigration Then Exit Sub

    Set pl = Tbl("tblPapers")
    If Not pl Is Nothing Then
        UnlockSheet pl.Parent
        For i = 1 To pl.ListRows.Count
            fam = Trim$(CStr(CellIn(pl, i, "Family").Value))
            If StrComp(fam, "Short Roll", vbTextCompare) = 0 Or StrComp(fam, "Long Roll", vbTextCompare) = 0 Then
                CellIn(pl, i, "Family").Value = "Roll"
            End If
        Next i
        RelockSheet pl.Parent
    End If

    UnlockSheet fl.Parent
    For i = 1 To fl.ListRows.Count
        fam = Trim$(CStr(CellIn(fl, i, "Family").Value))
        If StrComp(fam, "Short Roll", vbTextCompare) = 0 Or StrComp(fam, "Long Roll", vbTextCompare) = 0 Then
            If keepRow = 0 Then
                keepRow = i
            Else
                removeRow = i
            End If
        End If
    Next i
    If keepRow > 0 Then
        CellIn(fl, keepRow, "Family").Value = "Roll"
        CellIn(fl, keepRow, "Measurement basis").Value = "Roll"
    End If
    If removeRow > 0 Then fl.ListRows(removeRow).Delete
    RelockSheet fl.Parent
    Invalidate
End Sub

' -------------------------------------------------- student-supplied stock -
' "Supplied (Roll)"/"Supplied (Sheet)" are ordinary tblPapers rows, not
' synthetic in-memory stock - that is deliberate: every existing Reports/
' Summary Family/Type/Unit lookup and every dropdown-staging path already
' handles a real catalogue row with no further code, and Cost = 0 is what
' makes Paper Cost compute to zero (§5.1) with no change to any formula.
' Compatible() is the one place that needs to know a stock is one of these
' (clsStock.SuppliedByStudent) rather than checking its own size, since
' there isn't one until job entry.
Public Sub EnsureSuppliedByStudentColumn()
    Dim lo As ListObject, lc As ListColumn, ws As Worksheet
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    If Not ColumnExists(lo, "Supplied by student") Then
        UnlockSheet ws
        Set lc = lo.ListColumns.Add(ColIdx(lo, "Active") + 1)
        lc.Name = "Supplied by student"
        If Not lc.DataBodyRange Is Nothing Then
            StyleInputCell lc.DataBodyRange
            With lc.DataBodyRange.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = "Supplied by student"
                .InputMessage = "Yes only for the two 'Supplied' catalogue rows - paper cost always computes to zero for these regardless of Cost."
                .ErrorTitle = "Supplied by student"
                .ErrorMessage = "Choose Yes or No."
            End With
            lc.DataBodyRange.Value = "No"
        End If
        RelockSheet ws
    End If

    EnsureSuppliedStockRow "Supplied (Roll)", "Roll"
    EnsureSuppliedStockRow "Supplied (Sheet)", "Sheet"
End Sub

Private Sub EnsureSuppliedStockRow(ByVal Description As String, ByVal Family As String)
    Dim lo As ListObject, i As Long, r As ListRow, ws As Worksheet
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Description").Value)), Description, vbTextCompare) = 0 Then Exit Sub
    Next i

    UnlockSheet ws
    ' Always appended (ListRows.Add with no Position lands at the bottom) -
    ' StocksFor's dropdown order follows tblPapers row order (modLists.ApplyTo
    ' stages items in the order the Collection hands them over, itself
    ' mStocks.Keys' insertion order from LoadCatalog's own row scan), so
    ' this is what keeps both supplied options at the bottom of the Paper
    ' Stock list, as requested, with no extra sort/tie-break code needed.
    Set r = lo.ListRows.Add
    r.Range.Cells(1, ColIdx(lo, "StockID")).Value = "STK-SUP-" & UCase$(Family)
    r.Range.Cells(1, ColIdx(lo, "Description")).Value = Description
    r.Range.Cells(1, ColIdx(lo, "Paper type")).Value = "Student supplied"
    r.Range.Cells(1, ColIdx(lo, "Family")).Value = Family
    r.Range.Cells(1, ColIdx(lo, "Cost")).Value = 0
    r.Range.Cells(1, ColIdx(lo, "Active")).Value = "Yes"
    If ColumnExists(lo, "Supplied by student") Then
        r.Range.Cells(1, ColIdx(lo, "Supplied by student")).Value = "Yes"
    End If
    RelockSheet ws
    Invalidate
End Sub
