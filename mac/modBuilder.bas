Attribute VB_Name = "modBuilder"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' The VBA-project half of mac/build.py.
'
' build.ps1 reaches $wb.VBProject over COM. On a Mac nothing outside Excel
' can: AppleScript has no way into the VBA project, so these few calls have
' to run as VBA inside Excel, from a workbook that already holds them -
' mac/Builder.xlsm, made by hand once from this file (SETUP.md). That is all
' this module does: clear a project, add modules, fill them. Everything else
' (opening the copy, InitialiseWorkbook, StampBuild, Save As) is build.py.
'
' Builder.xlsm never needs remaking when this file changes. Every build
' loads THIS file afresh into its throwaway copy of Builder.xlsm, as module
' modBuilderLive, using whatever older modBuilder the copy already holds
' (BuilderAdd and BuilderAppend are all it needs), and then calls only
' modBuilderLive. Two consequences for editing this file:
'   - BuilderAdd and BuilderAppend must keep their names and arguments, since
'     an older Builder.xlsm uses them to load the newer code.
'   - No line may end in a " _" continuation. BuilderAppend inserts one line
'     at a time, and the VBE puts a blank line after every continuation line
'     inserted that way, which would break the statement.
'
' Code arrives as text from build.py rather than by VBComponents.Import:
' Excel for Mac is sandboxed, and VBA reading the .bas files out of the
' OneDrive folder would stop at a file-access prompt. Pasting is equivalent
' here because no module uses a non-default Attribute (ThisWorkbook is never
' imported anyway - build.ps1 pastes it too).
'
' Every function returns non-empty text. A macro AppleScript cannot find
' comes back empty with no error, so empty means "not found", never "done".
'
' Needs Excel > Preferences > Security > "Trust access to the VBA project
' object model".

Public Const BUILDER_VERSION As String = "2"

Private mBuf As String

Public Function BuilderReady() As String
    BuilderReady = "builder " & BUILDER_VERSION
End Function

' Every standard, class and form module removed, and ThisWorkbook emptied,
' as build.ps1's "clear any previously imported components" does.
Public Function BuilderClear(ByVal WbName As String) As String
    Dim vbp As Object, i As Long, n As Long
    On Error GoTo Fail
    Set vbp = Workbooks(WbName).VBProject
    For i = vbp.VBComponents.Count To 1 Step -1
        Select Case vbp.VBComponents(i).Type
            Case 1, 2, 3
                vbp.VBComponents.Remove vbp.VBComponents(i)
                n = n + 1
        End Select
    Next i
    EmptyModule vbp.VBComponents("ThisWorkbook").CodeModule
    BuilderClear = "ok " & n & " removed"
    Exit Function
Fail:
    BuilderClear = "ERROR " & Err.Number & ": " & Err.Description
End Function

' A new, EMPTY module. Kind 1 = standard, 2 = class. Emptied because the
' VBE's "Require Variable Declaration" option pre-fills Option Explicit, and
' the source carries its own.
Public Function BuilderAdd(ByVal WbName As String, ByVal ModName As String, ByVal Kind As Long) As String
    Dim c As Object
    On Error GoTo Fail
    Set c = Workbooks(WbName).VBProject.VBComponents.Add(Kind)
    c.Name = ModName
    EmptyModule c.CodeModule
    BuilderAdd = "ok"
    Exit Function
Fail:
    BuilderAdd = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Collects one piece of a module (lines separated by LF). Nothing is written
' until BuilderCommit: a module has to go in with ONE InsertLines call. Line
' by line, the VBE adds a blank line after every " _" continuation; in
' several pieces, it adds one at every join.
Public Function BuilderBuffer(ByVal Code As String) As String
    If Len(mBuf) > 0 Then mBuf = mBuf & vbLf
    mBuf = mBuf & Code
    BuilderBuffer = "ok " & Len(mBuf)
End Function

' Writes everything BuilderBuffer collected to the end of a module, in one
' call, and empties the buffer. InsertLines leaves a blank line after the
' text, which is removed again.
Public Function BuilderCommit(ByVal WbName As String, ByVal ModName As String) As String
    Dim cm As Object, n As Long
    On Error GoTo Fail
    Set cm = Workbooks(WbName).VBProject.VBComponents(ModName).CodeModule
    cm.InsertLines cm.CountOfLines + 1, mBuf
    n = cm.CountOfLines
    If n > 0 And Right$(mBuf, 1) <> vbLf Then
        If Len(cm.Lines(n, 1)) = 0 Then cm.DeleteLines n, 1
    End If
    mBuf = ""
    BuilderCommit = "ok " & cm.CountOfLines
    Exit Function
Fail:
    mBuf = ""
    BuilderCommit = "ERROR " & Err.Number & ": " & Err.Description & " (" & ModName & ")"
End Function

' Appends Code (lines separated by LF) one line at a time. Kept for loading
' this module into an older Builder.xlsm (see the top); BuilderBuffer and
' BuilderCommit are what builds use, since this mangles continuations.
Public Function BuilderAppend(ByVal WbName As String, ByVal ModName As String, ByVal Code As String) As String
    Dim cm As Object, lines() As String, i As Long, s As String
    On Error GoTo Fail
    Set cm = Workbooks(WbName).VBProject.VBComponents(ModName).CodeModule
    lines = Split(Code, vbLf)
    For i = LBound(lines) To UBound(lines)
        s = lines(i)
        If Len(s) > 0 Then If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
        cm.InsertLines cm.CountOfLines + 1, s
    Next i
    BuilderAppend = "ok " & cm.CountOfLines
    Exit Function
Fail:
    BuilderAppend = "ERROR " & Err.Number & ": " & Err.Description & " (" & ModName & ", line " & i & ")"
End Function

' Lines in a module, so build.py can check nothing was lost on the way.
Public Function BuilderLineCount(ByVal WbName As String, ByVal ModName As String) As String
    On Error GoTo Fail
    BuilderLineCount = CStr(Workbooks(WbName).VBProject.VBComponents(ModName).CodeModule.CountOfLines)
    Exit Function
Fail:
    BuilderLineCount = "ERROR " & Err.Number & ": " & Err.Description
End Function

Private Sub EmptyModule(ByVal cm As Object)
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
End Sub
