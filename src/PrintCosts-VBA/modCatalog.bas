Attribute VB_Name = "modCatalog"
Option Explicit

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

Public Sub Invalidate()
    mLoaded = False
End Sub

Public Sub LoadCatalog(Optional ByVal Force As Boolean = False)
    If mLoaded And Not Force Then Exit Sub

    Dim lo As ListObject, i As Long
    Dim s As clsStock, p As clsPrinterDef

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
            p.Families = CStr(CellIn(lo, i, "Supported families").Value)
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

' A stock is usable on a printer only when the stock's family is one the
' printer supports (spec 7.1, 10.5).
Public Function Compatible(ByVal Model As String, ByVal Description As String) As Boolean
    Dim p As clsPrinterDef, s As clsStock
    Set p = Prn(Model)
    Set s = Stock(Description)
    If Not p.Found Then Exit Function
    If Not s.Found Then Exit Function
    Compatible = InList(p.Families, s.Family)
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
