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

' The two built-in student-supplied stocks (see "student-supplied stock" at
' the foot of this module). Reserved names: not tblPapers rows.
Public Const SUPPLIED_ROLL As String = "Supplied (Roll)"
Public Const SUPPLIED_SHEET As String = "Supplied (Sheet)"

' tblPrinters[Colour mode] values (multi-pass costing, 0.11.0). Blank = single pass.
Public Const COLOUR_MODE_SINGLE As String = "single pass"
Public Const COLOUR_MODE_MULTI As String = "multi-pass"

Private mStocks As clsDict      ' by description
Private mPrinters As clsDict    ' by model
Private mTechs As clsDict       ' by name -> TechID
Private mDepts As clsDict       ' by department Name AND by each Alias -> clsDept
Private mColours As clsDict     ' by colour name -> clsColour (Consumables sheet)
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
    Set mDepts = New clsDict
    Set mColours = New clsDict

    Set lo = Tbl("tblPapers")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Description").Value) > 0 And Not IsBuiltInStock(CStr(CellIn(lo, i, "Description").Value)) Then
            Set s = New clsStock
            s.StockID = CStr(CellIn(lo, i, "StockID").Value)
            s.Description = CStr(CellIn(lo, i, "Description").Value)
            s.PaperType = CStr(CellIn(lo, i, "Paper type").Value)
            s.Measure = CStr(CellIn(lo, i, "Measure").Value)
            s.WidthMM = NumOf(CellIn(lo, i, "Width mm"))
            s.HeightMM = NumOf(CellIn(lo, i, "Height mm"))
            s.Cost = NumOf(CellIn(lo, i, "Cost"))
            s.Active = (StrComp(CStr(CellIn(lo, i, "Active").Value), "Yes", vbTextCompare) = 0)
            s.SuppliedByStudent = (StrComp(CStr(CellIn(lo, i, "Supplied by student").Value), "Yes", vbTextCompare) = 0)
            ' Paper the student supplies is never charged for, whatever the
            ' Cost cell says - the flag decides, not the number.
            If s.SuppliedByStudent Then s.Cost = 0
            s.Found = True
            mStocks.Add s.Description, s
        End If
    Next i
    ' Last, so both stay at the bottom of the Paper Stock dropdown.
    AddBuiltInStocks

    Set lo = Tbl("tblPrinters")
    For i = 1 To lo.ListRows.Count
        If Len(CellIn(lo, i, "Model").Value) > 0 Then
            Set p = New clsPrinterDef
            p.PrinterID = CStr(CellIn(lo, i, "PrinterID").Value)
            p.Model = CStr(CellIn(lo, i, "Model").Value)
            p.Consumable = CStr(CellIn(lo, i, "Consumable type").Value)
            p.RatePerM2 = NumOf(CellIn(lo, i, "Cost per m2"))
            ' Multi-pass (0.11.0): blank or absent reads as single pass / 0, so a
            ' printer row that predates the columns behaves exactly as before.
            If ColumnExists(lo, "Colour mode") Then
                p.MultiPass = (StrComp(Trim$(CStr(CellIn(lo, i, "Colour mode").Value)), COLOUR_MODE_MULTI, vbTextCompare) = 0)
            End If
            If ColumnExists(lo, "Template cost") Then p.TemplateCost = NumOf(CellIn(lo, i, "Template cost"))
            p.MaxRollWidthMM = NumOf(CellIn(lo, i, "Max roll width mm"))
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

    LoadDepartments
    LoadColours

    mLoaded = True
End Sub

' Colours (the Consumables sheet). Absent table = no colours, so a workbook
' that predates the sheet loads cleanly. Keyed by colour name, like printers
' and papers (design decision 26a); a name used twice goes to the first row.
Private Sub LoadColours()
    Dim lo As ListObject, i As Long, c As clsColour, nm As String
    Set lo = Tbl("tblColours")
    If lo Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        nm = Trim$(CStr(CellIn(lo, i, "Colour").Value))
        If Len(nm) > 0 Then
            If Not mColours.Exists(nm) Then
                Set c = New clsColour
                c.ColourID = CStr(CellIn(lo, i, "ColourID").Value)
                c.Colour = nm
                c.ConsumableType = Trim$(CStr(CellIn(lo, i, "Consumable type").Value))
                c.RatePerM2 = NumOf(CellIn(lo, i, "Cost per m2"))
                c.Active = (StrComp(Trim$(CStr(CellIn(lo, i, "Active").Value)), "Yes", vbTextCompare) = 0)
                c.Found = True
                mColours.Add nm, c
            End If
        End If
    Next i
End Sub

' Departments (the Departments sheet). Absent table = no departments, so a
' workbook that predates the sheet loads cleanly. A name or alias claimed by
' more than one row goes to the FIRST row that claims it (names before any
' alias of a later row); the sheet warns about the clash when it is typed.
Private Sub LoadDepartments()
    Dim lo As ListObject, i As Long, j As Long, d As clsDept, parts As Variant, a As String

    Set lo = Tbl("tblDepartments")
    If lo Is Nothing Then Exit Sub

    ' Names first, so a later row's alias can never shadow an earlier row's name.
    For i = 1 To lo.ListRows.Count
        If Len(Trim$(CStr(CellIn(lo, i, "Name").Value))) > 0 Then
            Set d = New clsDept
            d.DeptID = CStr(CellIn(lo, i, "DeptID").Value)
            d.Name = Trim$(CStr(CellIn(lo, i, "Name").Value))
            d.Free = (StrComp(Trim$(CStr(CellIn(lo, i, "Free").Value)), "Yes", vbTextCompare) = 0)
            d.Active = (StrComp(Trim$(CStr(CellIn(lo, i, "Active").Value)), "Yes", vbTextCompare) = 0)
            d.Found = True
            If Not mDepts.Exists(d.Name) Then mDepts.Add d.Name, d
        End If
    Next i
    For i = 1 To lo.ListRows.Count
        If Len(Trim$(CStr(CellIn(lo, i, "Name").Value))) > 0 Then
            Set d = mDepts.Obj(Trim$(CStr(CellIn(lo, i, "Name").Value)))
            If StrComp(d.DeptID, CStr(CellIn(lo, i, "DeptID").Value), vbBinaryCompare) = 0 Then
                parts = Split(CStr(CellIn(lo, i, "Aliases").Value), ";")
                For j = LBound(parts) To UBound(parts)
                    a = Trim$(CStr(parts(j)))
                    If Len(a) > 0 Then
                        If Not mDepts.Exists(a) Then mDepts.Add a, d
                    End If
                Next j
            End If
        End If
    Next i
End Sub

' ------------------------------------------------------- department lookup -
' THE one place a job is tied to a department. Today the tie is the job's
' Student Name matching a department's Name or an Alias (case-insensitive,
' trimmed; D19). A later release may give jobs a real Department column - this
' function is then the only thing that changes its source, and every consumer
' (the auto-disregard on entry, and the Department column on _Data that
' Reports and Summary read) keeps working unchanged.
'
' Includes inactive departments, so history stays classified; a Found = False
' object (never Nothing) means "not a department".
Public Function DepartmentFor(ByVal StudentName As String) As clsDept
    LoadCatalog
    StudentName = Trim$(StudentName)
    If Len(StudentName) > 0 Then
        If mDepts.Exists(StudentName) Then
            Set DepartmentFor = mDepts.Obj(StudentName)
            Exit Function
        End If
    End If
    Set DepartmentFor = New clsDept
End Function

' The department that should make a NEW job free: Active and Free = Yes.
' Found = False otherwise.
Public Function FreeDepartmentFor(ByVal StudentName As String) As clsDept
    Dim d As clsDept
    Set d = DepartmentFor(StudentName)
    If d.Found Then
        If d.Active And d.Free Then
            Set FreeDepartmentFor = d
            Exit Function
        End If
    End If
    Set FreeDepartmentFor = New clsDept
End Function

' Every department name the Reports filter offers: ALL defined departments,
' active or not, because old jobs for a deactivated department still need to
' be found. Sorted - a person reads the dropdown.
Public Function AllDepartments() As Collection
    Dim out As New Collection, lo As ListObject, i As Long, nm As String
    Set lo = Tbl("tblDepartments")
    If Not lo Is Nothing Then
        For i = 1 To lo.ListRows.Count
            nm = Trim$(CStr(CellIn(lo, i, "Name").Value))
            If Len(nm) > 0 Then out.Add nm
        Next i
    End If
    Set AllDepartments = SortedTextCollection(out)
End Function

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

' A colour by name; Found = False (never Nothing) when this workbook does not
' define it - an imported pass row may name a colour defined only elsewhere.
Public Function Colour(ByVal ColourName As String) As clsColour
    LoadCatalog
    If mColours.Exists(ColourName) Then
        Set Colour = mColours.Obj(ColourName)
    Else
        Set Colour = New clsColour
    End If
End Function

' Active colours of one consumable type, sorted - what a pass's Colour dropdown
' offers for a printer of that type (design decision 3).
Public Function ColoursFor(ByVal ConsumableType As String) As Collection
    Dim out As New Collection, k As Variant, c As clsColour
    LoadCatalog
    For Each k In mColours.Keys()
        Set c = mColours.Obj(CStr(k))
        If c.Active And StrComp(c.ConsumableType, Trim$(ConsumableType), vbTextCompare) = 0 Then out.Add c.Colour
    Next k
    Set ColoursFor = SortedTextCollection(out)
End Function

' The model name of the printer with this PrinterID, "" if none - how a job
' row's S_PrinterID is turned back into the printer it was stamped with.
Public Function PrinterModelById(ByVal PrinterID As String) As String
    Dim k As Variant, p As clsPrinterDef
    If Len(PrinterID) = 0 Then Exit Function
    LoadCatalog
    For Each k In mPrinters.Keys()
        Set p = mPrinters.Obj(CStr(k))
        If StrComp(p.PrinterID, PrinterID, vbBinaryCompare) = 0 Then
            PrinterModelById = p.Model
            Exit Function
        End If
    Next k
End Function

Public Function TechID(ByVal TechName As String) As String
    LoadCatalog
    TechID = CStr(mTechs.Item(TechName))
End Function

' A stock is usable on a printer only when the printer can take that kind of
' stock (roll/sheet, via Max roll width mm / Max sheet size) AND the stock's
' own size fits within that capacity. (Compatibility used to be decided by a
' paper "family" standing in for a width band - "Short Roll"/"Long Roll" -
' which was too coarse; the family concept has since been removed entirely,
' leaving Papers' Measure column - Sheet or Roll - as the one input.)
'
' The two built-in "Supplied (Roll)"/"Supplied (Sheet)" stocks (student-
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
        ' Multi-pass costing is sheet stock only (design decision 9).
        If p.MultiPass Then Exit Function
        If p.MaxRollWidthMM <= 0 Then Exit Function
        If s.PerJobSize Then
            Compatible = True
        Else
            Compatible = (s.WidthMM <= p.MaxRollWidthMM)
        End If
    ElseIf s.Measure = "Sheet" Then
        If Len(p.MaxSheetSize) = 0 Then Exit Function
        If s.PerJobSize Then
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
' a size fit rather than a family list.
Public Function CapacityText(ByVal p As clsPrinterDef) As String
    Dim parts As String
    If p.MaxRollWidthMM > 0 And Not p.MultiPass Then parts = "roll stock up to " & Format$(p.MaxRollWidthMM, "#,##0") & " mm wide"
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
' location whose capacity fits the given stock (spec 1a). Paper Stock can now be chosen before Printer, so Printer's own list
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
    DefaultSupplied lo, r.Range.Cells(1, 1).Row - lo.DataBodyRange.Row + 1
    DefaultFree lo, r.Range.Cells(1, 1).Row - lo.DataBodyRange.Row + 1
    ' Technicians, printers and papers each get their site-prefixed ID here,
    ' so it is there the moment the row is - see "catalogue IDs" below.
    FillCatalogId lo, TableName, r.Index
    RelockSheet ws
    Invalidate
    AppOn

    r.Range.Cells(1, 1).Select
End Sub

' Sets Active to "Yes" on table row RowNo when the table has an Active column
' and the cell is blank. Never overwrites an existing value. Caller has already
' unlocked the sheet.
Public Sub DefaultActive(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Active") Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, "Active").Value))) = 0 Then
        CellIn(lo, RowNo, "Active").Value = "Yes"
    End If
End Sub

' Papers table only: a blank "Supplied by student" becomes "No" - a stock is
' charged for unless it is explicitly marked as student-supplied. Caller has
' already unlocked the sheet.
Private Sub DefaultSupplied(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Supplied by student") Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, "Supplied by student").Value))) = 0 Then
        CellIn(lo, RowNo, "Supplied by student").Value = "No"
    End If
End Sub

' Departments table only: a blank Free becomes "No" - a listed department is
' charged like anyone else until it is explicitly marked Free. Caller has
' already unlocked the sheet.
Public Sub DefaultFree(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Free") Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, "Free").Value))) = 0 Then
        CellIn(lo, RowNo, "Free").Value = "No"
    End If
End Sub

' A stock marked Supplied by student never has a paper cost (LoadCatalog
' enforces it); this keeps the visible Cost cell honest as well. Caller has
' already unlocked the sheet.
Private Sub ZeroSuppliedCost(ByVal lo As ListObject, ByVal RowNo As Long)
    If Not ColumnExists(lo, "Supplied by student") Then Exit Sub
    If StrComp(Trim$(CStr(CellIn(lo, RowNo, "Supplied by student").Value)), "Yes", vbTextCompare) <> 0 Then Exit Sub
    If NumOf(CellIn(lo, RowNo, "Cost")) = 0 Then Exit Sub
    CellIn(lo, RowNo, "Cost").Value = 0
    Say "Paper supplied by the student is not charged for.", _
        "Cost has been set to 0 on '" & Trim$(CStr(CellIn(lo, RowNo, "Description").Value)) & "' because 'Supplied by student' is Yes.", _
        "Set 'Supplied by student' to No if the print room buys and charges for this stock."
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

        ' "Supplied (Roll)"/"Supplied (Sheet)" are built in (AddBuiltInStocks);
        ' a table row can not use either name.
        If IsBuiltInStock(CStr(CellIn(lo, n, "Description").Value)) Then
            Say "'" & Trim$(CStr(CellIn(lo, n, "Description").Value)) & "' is a built-in paper stock and can not be used as a name here.", _
                "The two 'Supplied' stocks are always available and are not stored on the Papers sheet.", _
                "Choose a different description. To record paper a student supplies in bulk, use your own name and set 'Supplied by student' to Yes."
            CellIn(lo, n, "Description").ClearContents
        End If

        If Len(Trim$(CStr(CellIn(lo, n, "Description").Value))) > 0 Then
            DefaultActive lo, n
            DefaultSupplied lo, n
            ZeroSuppliedCost lo, n
            FillCatalogId lo, "tblPapers", n
        End If
    Next c
    On Error GoTo 0
    RelockSheet ws
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    RelockSheet ws
    Err.Raise en, "OnPaperEdited", ed
End Sub

' ----------------------------------------------------- catalogue IDs ---
' Technicians, printers and papers each carry a synthetic ID (TechID,
' PrinterID, StockID) that nobody types by hand. It is allocated the way a
' Job ID is (modRegistry.NextJobId): the site (SET_SITE_ID), a table code and
' a four-digit number taken from a persisted high-water mark that only ever
' rises, e.g. MAIN-PRN-0001. The reason is the same too - two workbooks that
' each number from 1 would issue the same ID to different things, and a
' backup restore (modBackup.ApplyCatalogRows) matches rows by exactly this ID,
' so it would overwrite one site's printer with another's. With the site in
' the ID, rows from another workbook never collide with local ones and
' combining configurations appends instead of overwriting.
'
' Unlike a Job ID there is no per-location part - these tables are site-wide -
' so the counters live in tblSettings (TECH_ID_HWM, PRINTER_ID_HWM,
' STOCK_ID_HWM), not the per-location registry. They are marked "Read-only"
' in their Notes, which is what makes a restore skip them (a backup must
' never wind a counter back) and locks the cell against hand edits.
'
' As with NextJobId, a scan of the IDs currently in the table is only a FLOOR
' under the persisted value, never the source of truth: deleting the
' highest-numbered row lowers what the scan finds but never the counter.

' The four facts that differ between the three tables. False for any other
' table, so callers can pass a table name without checking it first.
Public Function CatalogIdSpec(ByVal TableName As String, ByRef IdHeader As String, ByRef Code As String, _
                              ByRef NameHeader As String, ByRef HwmKey As String) As Boolean
    Select Case TableName
        Case "tblTechnicians"
            IdHeader = "TechID": Code = "TCH": NameHeader = "Name": HwmKey = "TECH_ID_HWM"
        Case "tblPrinters"
            IdHeader = "PrinterID": Code = "PRN": NameHeader = "Model": HwmKey = "PRINTER_ID_HWM"
        Case "tblPapers"
            IdHeader = "StockID": Code = "STK": NameHeader = "Description": HwmKey = "STOCK_ID_HWM"
        Case "tblDepartments"
            IdHeader = "DeptID": Code = "DEP": NameHeader = "Name": HwmKey = "DEPT_ID_HWM"
        Case "tblColours"
            IdHeader = "ColourID": Code = "CLR": NameHeader = "Colour": HwmKey = "COLOUR_ID_HWM"
        Case Else
            Exit Function
    End Select
    CatalogIdSpec = True
End Function

' Sets a counter. Events are switched off around the write: it lands on the
' Settings sheet, whose Workbook_SheetChange would otherwise mark the whole
' catalogue dirty and rebind every dropdown for a change nobody made.
Private Sub SetCatalogHwm(ByVal HwmKey As String, ByVal Value As Long)
    Dim c As Range, ev As Boolean

    On Error Resume Next
    Set c = ThisWorkbook.Names("SET_" & HwmKey).RefersToRange
    On Error GoTo 0
    If c Is Nothing Then Exit Sub

    ev = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet c.Parent
    c.Value = Value
    RelockSheet c.Parent
    Application.EnableEvents = ev
End Sub

' The next ID for one catalogue table, advancing its counter.
Public Function NextCatalogId(ByVal lo As ListObject, ByVal TableName As String) As String
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String
    Dim prefix As String, hi As Long

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Function
    prefix = SettingText("SITE_ID", "SITE") & "-" & code & "-"

    hi = CLng(SettingNum(hwmKey, 0))
    hi = CLng(Application.WorksheetFunction.Max(hi, ScanMaxSuffix(lo, idHdr, prefix)))
    hi = hi + 1
    SetCatalogHwm hwmKey, hi
    NextCatalogId = prefix & Format$(hi, "0000")
End Function

' Gives table row RowNo its ID if it has none; never touches an existing one.
' Caller has already unlocked the sheet.
Public Sub FillCatalogId(ByVal lo As ListObject, ByVal TableName As String, ByVal RowNo As Long)
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Sub
    If Len(Trim$(CStr(CellIn(lo, RowNo, idHdr).Value))) > 0 Then Exit Sub
    CellIn(lo, RowNo, idHdr).Value = NextCatalogId(lo, TableName)
End Sub

' Raises a counter to cover every ID now in the table under this site's
' prefix - for after rows have been written in bulk (a restore), where the
' incoming numbers may exceed anything allocated locally so far. Never lowers
' it.
Public Sub SyncCatalogHwm(ByVal lo As ListObject, ByVal TableName As String)
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String
    Dim cur As Long, seen As Long

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Sub
    cur = CLng(SettingNum(hwmKey, 0))
    seen = ScanMaxSuffix(lo, idHdr, SettingText("SITE_ID", "SITE") & "-" & code & "-")
    If seen > cur Then SetCatalogHwm hwmKey, seen
End Sub

' Called from ThisWorkbook.Workbook_SheetChange for an edit on the Printers or
' Print Technicians sheet, with events already off (Papers has its own,
' OnPaperEdited). A row created by typing under the table never passes through
' AddCatalogRow, so this is what gives it an ID: once it has a name.
Public Sub OnCatalogEdited(ByVal TableName As String, ByVal Target As Range)
    Dim lo As ListObject, hit As Range, c As Range, ws As Worksheet, n As Long
    Dim idHdr As String, code As String, nameHdr As String, hwmKey As String
    Dim en As Long, ed As String, needsId As Boolean

    If Not CatalogIdSpec(TableName, idHdr, code, nameHdr, hwmKey) Then Exit Sub
    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set hit = Application.Intersect(Target, lo.DataBodyRange)
    If hit Is Nothing Then Exit Sub
    Set ws = lo.Parent

    ' Most edits need nothing from here (a price, a rename, a row deleted), and
    ' unlocking then re-locking the sheet for them is not free of side effects:
    ' it would re-protect a sheet the caller had deliberately left unlocked for
    ' a run of structural changes. So look first, and touch protection only when
    ' a named row really has no ID.
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If Len(Trim$(CStr(CellIn(lo, n, nameHdr).Value))) > 0 And _
           Len(Trim$(CStr(CellIn(lo, n, idHdr).Value))) = 0 Then
            needsId = True
            Exit For
        End If
    Next c
    If Not needsId Then Exit Sub

    UnlockSheet ws
    On Error GoTo Fail
    For Each c In hit.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If Len(Trim$(CStr(CellIn(lo, n, nameHdr).Value))) > 0 Then FillCatalogId lo, TableName, n
    Next c
    On Error GoTo 0
    RelockSheet ws
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    RelockSheet ws
    Err.Raise en, "OnCatalogEdited", ed
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

' ------------------------------------------------------ catalogue clear ---
' "Clear table" on Print Technicians, Printers and Papers (direct user
' request, 2026-09-29): empties ONE catalogue table so a bureau can start that
' list afresh without touching the others - e.g. a new academic year with new
' technicians but the same printers. Only these three tables have the button;
' the Settings lookup tables do not.
'
' ClearCatalogTable asks first (naming the table, how many rows and which,
' and what the consequences are - the same "never a bare are you sure" rule
' RemoveCatalogRow and modJobs.ClearAll follow). DoClearCatalogTable is the
' unprompted work, split out because Ask always declines in quiet mode, so a
' test script could otherwise never exercise the clear itself.
'
' Recorded jobs are not touched: every job snapshots the price it was costed
' at (the S_* columns) - the same guarantee RemoveCatalogRow gives. What does
' change is every print room's dropdowns (Invalidate marks them stale; they
' rebind on the next sheet change).
Public Sub ClearCatalogTable(ByVal TableName As String)
    Dim lo As ListObject, n As Long, i As Long, label As String, keyHdr As String
    Dim listing As String, shown As Long, extra As String, msg As String

    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub

    ClearTableInfo TableName, label, keyHdr, extra
    n = RowCount(lo)
    If n = 0 Then
        Say "There is nothing to clear.", "The " & label & " table has no rows."
        Exit Sub
    End If

    ' Name the first few rows so the person can see WHAT is about to go,
    ' not just how many.
    For i = 1 To lo.ListRows.Count
        If Len(Trim$(CStr(CellIn(lo, i, keyHdr).Value))) > 0 Then
            If shown < 5 Then
                listing = listing & "   - " & Trim$(CStr(CellIn(lo, i, keyHdr).Value)) & vbCrLf
            End If
            shown = shown + 1
        End If
    Next i
    If shown > 5 Then listing = listing & "   ...and " & (shown - 5) & " more" & vbCrLf

    msg = "Clear table - " & label & vbCrLf & vbCrLf & _
        "This will permanently delete all " & n & " row" & IIf(n = 1, "", "s") & " from the " & label & " table:" & vbCrLf & _
        listing & vbCrLf & _
        "They will disappear from every print room's dropdowns straight away." & vbCrLf & _
        "Print jobs already recorded keep their frozen prices and are not affected." & vbCrLf & _
        extra & _
        "The other configuration sheets are not affected." & vbCrLf & vbCrLf & _
        "This cannot be undone - use Backup workbook first if you may want these back. Continue?"

    If Not Ask(msg, "Clear table") Then Exit Sub

    DoClearCatalogTable TableName
    Say n & " row" & IIf(n = 1, "", "s") & " deleted from the " & label & " table.", _
        "A note of what was removed has been kept in the workbook's audit log."
End Sub

' What the confirmation calls the table, which column names its rows, and any
' extra sentence that only applies to this table (ends with vbCrLf if set).
Private Sub ClearTableInfo(ByVal TableName As String, ByRef Label As String, ByRef KeyHdr As String, ByRef Extra As String)
    Select Case TableName
        Case "tblTechnicians"
            Label = "Print Technicians": KeyHdr = "Name"
        Case "tblPrinters"
            Label = "Printers": KeyHdr = "Model"
            Extra = "Each print room's 'Select printers' choice refers to printers by name, so rooms will need their printers re-selected once you add new ones." & vbCrLf
        Case "tblDepartments"
            Label = "Departments": KeyHdr = "Name"
            Extra = "Jobs already recorded for these departments stay as they are, but Reports and Summary can no longer group them under a department." & vbCrLf
        Case "tblColours"
            Label = "Consumables (colours)": KeyHdr = "Colour"
            Extra = "Pass rows already recorded keep their stamped ink rate, but their colour will no longer be found here." & vbCrLf
        Case "tblPapers"
            Label = "Papers": KeyHdr = "Description"
            Extra = "The built-in 'Supplied (Roll)' and 'Supplied (Sheet)' stocks are not stored here and stay available." & vbCrLf
        Case Else
            Err.Raise vbObjectError + 2101, "ClearCatalogTable", "Clear table is not available for '" & TableName & "'."
    End Select
End Sub

' The unprompted clear. Leaves exactly one blank row rather than zero rows:
' with DataBodyRange gone there is no row left for Excel to copy formatting
' and validation from, so a later ListRows.Add would produce a bare row - the
' same reason modJobs.ClearAll keeps row 1. AddCatalogRow already reuses that
' blank row (Count = 1 And IsBlankRow) and RowCount reports it as zero rows.
' Cells holding a formula are skipped rather than cleared, so a calculated
' column survives.
Public Sub DoClearCatalogTable(ByVal TableName As String)
    Dim lo As ListObject, ws As Worksheet, n As Long, i As Long, c As Range
    Dim label As String, keyHdr As String, extra As String

    Set lo = Tbl(TableName)
    If lo Is Nothing Then Exit Sub
    ClearTableInfo TableName, label, keyHdr, extra
    Set ws = lo.Parent
    n = RowCount(lo)
    If n = 0 Then Exit Sub

    AppOff
    LogAudit "Clear table", label, n & " row" & IIf(n = 1, "", "s") & " deleted"
    UnlockSheet ws
    For i = lo.ListRows.Count To 2 Step -1
        lo.ListRows(i).Delete
    Next i
    For Each c In lo.ListRows(1).Range.Cells
        If Not c.HasFormula Then c.ClearContents
    Next c
    RelockSheet ws
    Invalidate
    AppOn
End Sub

' -------------------------------------------------- student-supplied stock -
' Two kinds of "supplied by the student" paper:
'
' 1. The built-in "Supplied (Roll)" / "Supplied (Sheet)" stocks. They are NOT
'    rows of tblPapers: they exist for every printer and location alike, so
'    LoadCatalog adds them in code (AddBuiltInStocks) and nobody can delete,
'    rename or re-cost them. The two names are reserved (IsBuiltInStock,
'    OnPaperEdited). They have no catalogue size - the real size is entered
'    on each job (Print Width mm / Sheet size) - and no paper cost.
' 2. Any tblPapers row with Supplied by student = Yes, e.g. a bulk delivery
'    of paper a student supplies for many jobs. It defaults to No, keeps its
'    own catalogue size like any other stock, and LoadCatalog forces its
'    Cost to 0.

Public Function IsBuiltInStock(ByVal Description As String) As Boolean
    Description = Trim$(Description)
    IsBuiltInStock = (StrComp(Description, SUPPLIED_ROLL, vbTextCompare) = 0) _
                  Or (StrComp(Description, SUPPLIED_SHEET, vbTextCompare) = 0)
End Function

Private Sub AddBuiltInStocks()
    AddBuiltInStock "STK-SUP-ROLL", SUPPLIED_ROLL, "Roll"
    AddBuiltInStock "STK-SUP-SHEET", SUPPLIED_SHEET, "Sheet"
End Sub

Private Sub AddBuiltInStock(ByVal StockID As String, ByVal Description As String, ByVal Measure As String)
    Dim s As New clsStock
    s.StockID = StockID
    s.Description = Description
    s.PaperType = "Student supplied"
    s.Measure = Measure             ' "Roll" / "Sheet" - fixed, not looked up
    s.Cost = 0
    s.Active = True
    s.SuppliedByStudent = True
    s.PerJobSize = True
    s.Found = True
    mStocks.Add s.Description, s
End Sub

' The Yes/No dropdown and its help text on tblPapers[Supplied by student].
' Re-applied here so the wording matches the behaviour above rather than
' whatever the .xlsx template happens to carry.
Public Sub EnsureSuppliedColumnValidation()
    Dim lo As ListObject, ws As Worksheet
    Set lo = Tbl("tblPapers")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set ws = lo.Parent

    UnlockSheet ws
    With lo.ListColumns("Supplied by student").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Supplied by student"
        .InputMessage = "Yes for paper the student supplies, e.g. a bulk delivery used across many jobs: paper cost is always zero, whatever Cost says. No (the default) for stock the print room buys and charges for."
        .ErrorTitle = "Supplied by student"
        .ErrorMessage = "Choose Yes or No."
    End With
    RelockSheet ws
End Sub
