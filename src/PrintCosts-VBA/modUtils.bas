Attribute VB_Name = "modUtils"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' Shared plumbing: application state, error reporting, table and name access.

Public Const MARKER As String = "PRINTLOC/v1"
Public Const MARKER_CELL As String = "AZ1"
' 1.1 (2026-09-22, snag list item 1c): added the Paid job-row column.
' 1.2 (printer/paper compatibility rework): added the "Sheet size" job-row
' column (the student-supplied-stock size override for sheet jobs, mirroring
' Print Width mm's existing role for roll jobs). Second schema bump since
' inception - every other 0.x.x release left the job-row shape untouched
' (§3.5: reordering or adding a registry/settings column never counts, only
' a genuine job-row column does).
Public Const SCHEMA_VER As String = "1.2"
Public Const JOBS_PREFIX As String = "tblJobs_"
Public Const LIST_SEP As String = ";"

Private mCalc As XlCalculation
Private mDepth As Long

' Quiet mode. When a build script or an automated test drives the workbook
' over COM there is nobody to dismiss a dialog, and a single MsgBox would hang
' the run indefinitely. In quiet mode messages are collected instead of shown
' and can be read back afterwards, and Ask answers No - an unattended run must
' never confirm a destructive operation on the user's behalf.
Public gQuiet As Boolean
Private mLog As String

' ---------------------------------------------------------------- app state ---
' Every public entry point brackets its work with these. Leaving EnableEvents
' False after an error would silently break the workbook for the rest of the
' session, so the restore belongs in the error path too, not just the happy one.
Public Sub AppOff()
    If mDepth = 0 Then
        mCalc = Application.Calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If
    mDepth = mDepth + 1
End Sub

Public Sub AppOn()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth = 0 Then
        Application.Calculation = mCalc
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub

Public Sub AppReset()
    mDepth = 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' --------------------------------------------------------------- messaging ---
' Spec section 15 wants three things from every message: what is wrong, why,
' and what to do about it. Say() takes them as separate arguments so that a
' caller cannot quietly omit the third.
Public Sub Say(ByVal Problem As String, Optional ByVal Reason As String = "", Optional ByVal Remedy As String = "", Optional ByVal Title As String = "Print log")
    Dim s As String
    s = Problem
    If Len(Reason) > 0 Then s = s & vbCrLf & vbCrLf & Reason
    If Len(Remedy) > 0 Then s = s & vbCrLf & vbCrLf & Remedy
    If gQuiet Then
        Note Title & ": " & s
    Else
        MsgBox s, vbInformation, Title
    End If
End Sub

Public Function Ask(ByVal Prompt As String, Optional ByVal Title As String = "Please confirm") As Boolean
    If gQuiet Then
        Note "DECLINED (quiet mode) - " & Title & ": " & Prompt
        Ask = False
        Exit Function
    End If
    Ask = (MsgBox(Prompt, vbExclamation + vbYesNo + vbDefaultButton2, Title) = vbYes)
End Function

Public Sub ReportError(ByVal Where As String)
    Dim s As String
    AppReset
    s = "Something went wrong in " & Where & "." & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "No changes have been left half-applied. If this keeps happening, " & "note what you were doing and pass it on with this message."
    If gQuiet Then
        Note "ERROR - " & s
    Else
        MsgBox s, vbCritical, "Print log"
    End If
End Sub

' ------------------------------------------------------- quiet-mode control ---
' Called from a build or test script over COM, never from the workbook itself.
Public Sub SetQuiet(ByVal Quiet As Boolean)
    gQuiet = Quiet
    mLog = ""
End Sub

Public Function QuietLog() As String
    QuietLog = mLog
End Function

Private Sub Note(ByVal s As String)
    mLog = mLog & s & vbCrLf & String$(60, "-") & vbCrLf
End Sub

' ------------------------------------------------------------------ tables ---
Public Function JobsTable(ByVal ws As Worksheet) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If Left$(lo.Name, Len(JOBS_PREFIX)) = JOBS_PREFIX Then
            Set JobsTable = lo
            Exit Function
        End If
    Next lo
End Function

Public Function Tbl(ByVal TableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, TableName, vbTextCompare) = 0 Then
                Set Tbl = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Column index by header text. Nothing in this project addresses a column by
' number, so inserting one later cannot silently shift the code.
Public Function ColIdx(ByVal lo As ListObject, ByVal Header As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, Header, vbTextCompare) = 0 Then
            ColIdx = i
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 513, "ColIdx", "Column '" & Header & "' was not found in table '" & lo.Name & "'."
End Function

Public Function CellIn(ByVal lo As ListObject, ByVal RowNo As Long, ByVal Header As String) As Range
    Set CellIn = lo.ListRows(RowNo).Range.Cells(1, ColIdx(lo, Header))
End Function

' Whether a column exists, without ColIdx's raise-if-missing behaviour -
' for code that reads an optional/newly-added column (the printer capacity
' columns, "Supplied by student", "Sheet size") and must not fail against a
' workbook that hasn't had the matching Ensure* migration run yet.
Public Function ColumnExists(ByVal lo As ListObject, ByVal Header As String) As Boolean
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, Header, vbTextCompare) = 0 Then
            ColumnExists = True
            Exit Function
        End If
    Next i
End Function

' Range.ColumnWidth is in "characters of the Normal-style font" (Calibri 11
' throughout this workbook, per xl/styles.xml), not pixels, and Excel itself
' is the only thing that draws the relationship between the two - there is no
' Range.WidthInPixels to set directly. Calibri 11's own metric (7px per
' character plus Excel's fixed 5px of cell padding: pixels = ColumnWidth*7+5)
' is the same constant openpyxl/xlsxwriter use and matches Excel's own
' column-width dialog for this font. Always rounds DOWN, so a "no more than
' Npx" width can never render one pixel over - callers asking for an exact
' target (not a cap) accept the same sub-pixel rounding for simplicity.
Public Function ColWidthForPx(ByVal px As Double) As Double
    ColWidthForPx = Int(((px - 5) / 7) * 100) / 100
End Function

Public Function RowCount(ByVal lo As ListObject) As Long
    If lo.ListRows.Count = 0 Then
        RowCount = 0
    ElseIf IsBlankRow(lo, 1) And lo.ListRows.Count = 1 Then
        RowCount = 0
    Else
        RowCount = lo.ListRows.Count
    End If
End Function

' A row that does not exist is not a blank row - False, not an error. Every
' call site in this project (modJobs, modImport, modCatalog, modSnapshot,
' modBackup) guards this with `lo.ListRows.Count = 1 And IsBlankRow(lo, 1)`,
' and VBA's And does NOT short-circuit: both operands are evaluated
' regardless of the first, so IsBlankRow(lo, 1) still ran - and raised
' "Subscript out of range" - even when Count was 0 and the left operand had
' already failed. Dormant for years because nothing in ordinary use ever
' left a table at genuinely zero rows (ClearAll and friends always leave the
' one blank templated row); found 2026-09-22 restoring a workbook backup
' into a catalogue table emptied down to zero rows by hand (modBackup,
' snag 4a) - the first code path in this project to actually hit that state.
Public Function IsBlankRow(ByVal lo As ListObject, ByVal RowNo As Long) As Boolean
    If RowNo < 1 Or RowNo > lo.ListRows.Count Then Exit Function
    IsBlankRow = (Application.WorksheetFunction.CountA(lo.ListRows(RowNo).Range) = 0)
End Function

' ------------------------------------------------------------------- names ---
Public Function LocRange(ByVal ws As Worksheet, ByVal RefName As String) As Range
    On Error GoTo Missing
    Set LocRange = ws.Names(RefName).RefersToRange
    Exit Function
Missing:
    Set LocRange = Nothing
End Function

Public Function LocValue(ByVal ws As Worksheet, ByVal RefName As String) As String
    Dim r As Range
    Set r = LocRange(ws, RefName)
    If Not r Is Nothing Then LocValue = CStr(r.Value)
End Function

' A sheet is a print room if its marker cell says so.
'
' This originally looked for a sheet-scoped defined name, LOC_Marker. The cell
' is the better test: it copies with a duplicated sheet exactly as a scoped
' name does, it cannot be silently renamed by Excel, and it is one moving part
' rather than two.
Public Function IsLocation(ByVal ws As Worksheet) As Boolean
    On Error GoTo No
    IsLocation = (StrComp(Trim$(CStr(ws.Range(MARKER_CELL).Value)), MARKER, vbTextCompare) = 0)
    Exit Function
No:
    IsLocation = False
End Function

' ------------------------------------------------------------------ pieces ---
Public Function SplitList(ByVal s As String) As Variant
    If Len(Trim$(s)) = 0 Then
        SplitList = Array()
    Else
        SplitList = Split(s, LIST_SEP)
    End If
End Function

Public Function InList(ByVal s As String, ByVal Item As String) As Boolean
    Dim p As Variant, i As Long
    p = SplitList(s)
    For i = LBound(p) To UBound(p)
        If StrComp(Trim$(CStr(p(i))), Trim$(Item), vbTextCompare) = 0 Then
            InList = True
            Exit Function
        End If
    Next i
End Function

' A date cell's serial number, or 0 if it does not hold one.
'
' Not Val(). Val() takes a String, so passing it a date cell coerces the value
' to text first and then reads the leading digits - "15/09/2026 10:24" comes
' back as 15, a date in January 1900. .Value2 hands over the underlying serial
' without going through formatting at all.
Public Function DateSerialOf(ByVal c As Range) As Double
    Dim v As Variant
    v = c.Value2
    If IsNumeric(v) Then
        If Len(Trim$(CStr(v))) > 0 Then DateSerialOf = CDbl(v)
    End If
End Function

' A numeric cell's value, or 0 if it does not hold one.
'
' Not Val(), for exactly the reason DateSerialOf is not Val(). Val takes a
' String, so a Double is coerced to text using the MACHINE'S decimal separator
' and then parsed back expecting a US point. On any comma-decimal locale
' Val(0.31) is Val("0,31") = 0 - and 0.31 here is a paper cost, a consumable
' rate or a stock dimension, so every price in the workbook silently becomes
' zero and every cost computes to nothing. Nothing raises; the figures are
' just wrong.
'
' Design 8.2 states this rule for date cells because that is where it was
' first caught. The cause is the round-trip through a locale-formatted string,
' which applies to any number with a fractional part. .Value2 hands over the
' underlying number without going near formatting.
Public Function NumOf(ByVal c As Range) As Double
    Dim v As Variant
    v = c.Value2
    If IsNumeric(v) Then
        If Len(Trim$(CStr(v))) > 0 Then NumOf = CDbl(v)
    End If
End Function

Public Function CurrentUser() As String
    CurrentUser = Application.UserName
    If Len(Trim$(CurrentUser)) = 0 Then CurrentUser = "unknown"
End Function

' A text-sorted copy of a Collection. An insertion sort is plenty here - every
' caller is a catalogue list numbering in the tens, not thousands.
Public Function SortedTextCollection(ByVal src As Collection) As Collection
    Dim vals() As String, n As Long, i As Long, j As Long, tmp As String
    Dim out As New Collection

    n = src.Count
    If n = 0 Then
        Set SortedTextCollection = out
        Exit Function
    End If

    ReDim vals(1 To n)
    For i = 1 To n
        vals(i) = CStr(src(i))
    Next i
    For i = 2 To n
        tmp = vals(i)
        j = i - 1
        Do While j >= 1
            If StrComp(vals(j), tmp, vbTextCompare) <= 0 Then Exit Do
            vals(j + 1) = vals(j)
            j = j - 1
        Loop
        vals(j + 1) = tmp
    Next i
    For i = 1 To n
        out.Add vals(i)
    Next i
    Set SortedTextCollection = out
End Function
