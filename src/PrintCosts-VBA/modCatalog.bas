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
    ' ListRows.Add copies formatting and validation but no VALUES, so Active
    ' would be blank - and LoadCatalog reads only a literal "Yes" as active, so
    ' a row left blank is silently excluded from every dropdown.
    DefaultActive lo, r.Range.Cells(1, 1).Row - lo.DataBodyRange.Row + 1
    RelockSheet ws
    Invalidate
    AppOn

    r.Range.Cells(1, 1).Select
End Sub

' Sets Active to "Yes" on table row RowNo when the table has an Active column
' and the cell is blank. Never overwrites an existing value. Caller has already
' unlocked the sheet.
Private Sub DefaultActive(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Active") Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, "Active").Value))) = 0 Then
        CellIn(lo, RowNo, "Active").Value = "Yes"
    End If
End Sub

' Called from ThisWorkbook.Workbook_SheetChange for every edit on the Papers
' sheet, with events already off (so the writes below do not re-enter it).
'
' 1. Size mode "Standard": fills Width mm / Height mm from the Std. size chosen,
'    as the note on the Papers sheet promises. Compatible() compares these two
'    numbers with the printer's capacity, so a sheet stock left at 0 x 0 fits
'    no printer and never reaches a dropdown. Fires only when Size mode or Std.
'    size was the cell edited, so a later hand-edit of Width/Height is kept.
' 2. Active: defaults to "Yes" on a row that has a Description but no Active
'    value - covers a row created by typing under the table (Excel auto-extend),
'    which never passes through AddCatalogRow.
Public Sub OnPaperEdited(ByVal Target As Range)
    Dim lo As ListObject, hit As Range, c As Range, ws As Worksheet
    Dim n As Long, hdr As String, sizeName As String
    Dim w As Double, h As Double, en As Long, ed As String

    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set hit = Application.Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub
    Set ws = lo.Parent

    UnlockSheet ws
    On Error GoTo Fail
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        hdr = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)

        If StrComp(hdr, "Size mode", vbTextCompare) = 0 Or StrComp(hdr, "Std. size", vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(CellIn(lo, n, "Size mode").Value)), "Standard", vbTextCompare) = 0 Then
                sizeName = Trim$(CStr(CellIn(lo, n, "Std. size").Value))
                If Len(sizeName) > 0 Then
                    ' Plain local Doubles, not fields of an object - see the
                    ' ByRef note in LoadCatalog.
                    If StdSizeDims(sizeName, w, h) Then
                        CellIn(lo, n, "Width mm").Value = w
                        CellIn(lo, n, "Height mm").Value = h
                    End If
                End If
            End If
        End If

        If Len(Trim$(CStr(CellIn(lo, n, "Description").Value))) > 0 Then DefaultActive lo, n
    Next c
    On Error GoTo 0
    RelockSheet ws
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    RelockSheet ws
    Err.Raise en, "OnPaperEdited", ed
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

' ---------------------------------------------------- standard sizes name ---
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

' -------------------------------------------------- student-supplied stock -
' "Supplied (Roll)"/"Supplied (Sheet)" ship as ordinary tblPapers rows (Cost 0,
' Supplied by student = Yes). Setup only restores one if it has been deleted:
' an existing row is never touched. Appended at the bottom so the Paper Stock
' dropdown keeps both supplied options last (StocksFor follows row order).
Public Sub EnsureSuppliedStockRows()
    RestoreSuppliedStockRow "Supplied (Roll)", "Roll"
    RestoreSuppliedStockRow "Supplied (Sheet)", "Sheet"
End Sub

Private Sub RestoreSuppliedStockRow(ByVal Description As String, ByVal Family As String)
    Dim lo As ListObject, i As Long, r As ListRow, ws As Worksheet
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(CellIn(lo, i, "Description").Value)), Description, vbTextCompare) = 0 Then Exit Sub
    Next i

    UnlockSheet ws
    Set r = lo.ListRows.Add
    r.Range.Cells(1, ColIdx(lo, "StockID")).Value = "STK-SUP-" & UCase$(Family)
    r.Range.Cells(1, ColIdx(lo, "Description")).Value = Description
    r.Range.Cells(1, ColIdx(lo, "Paper type")).Value = "Student supplied"
    r.Range.Cells(1, ColIdx(lo, "Family")).Value = Family
    r.Range.Cells(1, ColIdx(lo, "Cost")).Value = 0
    r.Range.Cells(1, ColIdx(lo, "Active")).Value = "Yes"
    r.Range.Cells(1, ColIdx(lo, "Supplied by student")).Value = "Yes"
    RelockSheet ws
    Invalidate
End Sub