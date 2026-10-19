Attribute VB_Name = "modScript"
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' A text-only way in for scripts that cannot hand VBA an object.
'
' Over COM (build.ps1, the test-*.ps1 scripts) a caller passes a Worksheet,
' Range or ListObject straight to a macro. AppleScript - the only way a
' script reaches Excel for Mac (mac/xlmac.py) - can pass text, numbers and
' True/False, nothing else. ScriptRun takes a macro's arguments as text
' tokens instead, turns each into its object, and makes the call:
'
'   @@wb                   ThisWorkbook
'   @@sheet:Name           ThisWorkbook.Worksheets(Name)
'   @@table:Name           the ListObject called Name, on whichever sheet
'   @@range:Sheet!A1:B2    that range (the sheet name unquoted)
'   @@spill:Sheet!H11      that cell's SpillingToRange
'   @@shape:Sheet!Name     the Form Controls object behind that shape (a
'                          Button, a DropDown), for ScriptProp
'   @@h:3                  an object or array an earlier call returned
'
' Anything else is passed through as it came. The result comes back as
' text: "ok:" and the value, or "ok:@@h:N" for an object or array, which is
' kept here so a later call can hand it back (ReadImportRows' rows into
' ApplyImportConfirmed, for example). "ERROR n: description" if it fails.
' Never empty, because AppleScript reports a macro it cannot find as an
' empty result.
'
' The arguments arrive packed into ONE string, Args, rather than as
' optional parameters: on Excel for Mac 16.114, AppleScript calling a
' function that has several optional parameters it leaves out crashes the
' VBA engine outright (EXC_BAD_ACCESS), so this signature has none. Each
' argument is a type letter and its text - s (text), n (number, "." as the
' decimal point), b (1 or 0) - and they are separated by Chr(31), the ASCII
' unit separator, which no cell value or name contains. mac/xlmac.py's
' Workbook.call does the packing. The call onward to Macro is VBA to VBA,
' where optional parameters are fine.
'
' Nothing in the workbook calls this, and it adds no capability: it only
' reaches routines that are Public already, the same ones a COM caller can
' run. Handles live until ScriptReset or the workbook closes.

Private Const TOK As String = "@@"
Private Const MAX_ARGS As Long = 8
Private Const SEP As Long = 31

Private mHandles As Collection

Public Function ScriptRun(ByVal Macro As String, ByVal Args As String) As String
    Dim parts() As String, a(1 To MAX_ARGS) As Variant
    Dim n As Long, i As Long, full As String, r As Variant

    On Error GoTo Fail
    If Len(Args) > 0 Then
        parts = Split(Args, Chr$(SEP))
        n = UBound(parts) - LBound(parts) + 1
        If n > MAX_ARGS Then Err.Raise vbObjectError + 6, "ScriptRun", "At most " & MAX_ARGS & " arguments."
        For i = 1 To n
            Resolve Unpack(parts(LBound(parts) + i - 1)), a(i)
        Next i
    End If

    If Not MacroExists(Macro) Then Err.Raise vbObjectError + 8, "ScriptRun", "No macro called '" & Macro & "'."
    full = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & Macro
    Select Case n
        Case 0: Keep r, Application.Run(full)
        Case 1: Keep r, Application.Run(full, a(1))
        Case 2: Keep r, Application.Run(full, a(1), a(2))
        Case 3: Keep r, Application.Run(full, a(1), a(2), a(3))
        Case 4: Keep r, Application.Run(full, a(1), a(2), a(3), a(4))
        Case 5: Keep r, Application.Run(full, a(1), a(2), a(3), a(4), a(5))
        Case 6: Keep r, Application.Run(full, a(1), a(2), a(3), a(4), a(5), a(6))
        Case 7: Keep r, Application.Run(full, a(1), a(2), a(3), a(4), a(5), a(6), a(7))
        Case 8: Keep r, Application.Run(full, a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8))
    End Select

    If IsObject(r) Or IsArray(r) Then
        ScriptRun = "ok:" & AddHandle(r)
    ElseIf IsEmpty(r) Or IsNull(r) Then
        ScriptRun = "ok:"
    ElseIf IsError(r) Then
        ScriptRun = "ok:#ERROR"
    Else
        ScriptRun = "ok:" & CStr(r)
    End If
    Exit Function
Fail:
    ScriptRun = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Drops every kept handle.
Public Function ScriptReset() As String
    Set mHandles = Nothing
    ScriptReset = "ok:"
End Function

' The value of something a handle or token holds, for a script that needs to
' look inside it: Prop is a property name - dotted ("Rows.Count") and
' indexed ("List 2") both work - or, for an array, "item r c" / "item r" /
' "ubound".
Public Function ScriptProp(ByVal Handle As String, ByVal Prop As String) As String
    Dim v As Variant, p() As String
    Dim chain() As String, o As Object, i As Long, last As String, r As Variant
    On Error GoTo Fail
    Resolve Handle, v
    p = Split(Prop, " ")
    ' IsObject first: IsArray on a multi-cell Range is True (it looks through
    ' to the Range's Value), which would send a range down the array branch.
    If Not IsObject(v) And IsArray(v) Then
        Select Case p(0)
            Case "ubound": ScriptProp = "ok:" & UBound(v, 1)
            Case "item"
                If UBound(p) >= 2 Then
                    ScriptProp = "ok:" & CStr(v(CLng(p(1)), CLng(p(2))))
                Else
                    ScriptProp = "ok:" & CStr(v(CLng(p(1))))
                End If
            Case Else: Err.Raise vbObjectError + 1, "ScriptProp", "Unknown array query '" & Prop & "'."
        End Select
        Exit Function
    End If
    ' Any property by name, read with CallByName: a dotted chain walks objects
    ' ("Rows.Count"), and a number after a space indexes the last one
    ' ("List 2" is DropDown.List(2)). Scalars only - the last step must not
    ' return an object.
    chain = Split(p(0), ".")
    Set o = v
    For i = LBound(chain) To UBound(chain) - 1
        Set o = CallByName(o, chain(i), VbGet)
    Next i
    last = chain(UBound(chain))
    If UBound(p) >= 1 Then
        r = CallByName(o, last, VbGet, CLng(p(1)))
    Else
        r = CallByName(o, last, VbGet)
    End If
    ScriptProp = "ok:" & CStr(r)
    Exit Function
Fail:
    ScriptProp = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Whether Macro ("Proc" or "module.Proc") is a procedure in this project.
' Checked before Application.Run because on Excel for Mac 16.114 running a
' name that does not exist crashes Excel, where Windows raises error 1004.
' Needs "Trust access to the VBA project object model"; without it the
' project cannot be read, and this says yes and leaves it to Run.
Private Function MacroExists(ByVal Macro As String) As Boolean
    Dim comps As Object, c As Object, modName As String, proc As String, k As Long

    On Error Resume Next
    Set comps = ThisWorkbook.VBProject.VBComponents
    If Err.Number <> 0 Or comps Is Nothing Then
        MacroExists = True
        Exit Function
    End If
    On Error GoTo 0

    k = InStr(Macro, ".")
    If k > 0 Then
        modName = Left$(Macro, k - 1)
        proc = Mid$(Macro, k + 1)
    Else
        proc = Macro
    End If
    For Each c In comps
        If Len(modName) = 0 Or StrComp(c.Name, modName, vbTextCompare) = 0 Then
            If HasProc(c.CodeModule, proc) Then
                MacroExists = True
                Exit Function
            End If
        End If
    Next c
End Function

' ProcStartLine raises an error for a name the module does not hold. Kind 0
' is a Sub or Function; properties are not run through Application.Run.
Private Function HasProc(ByVal cm As Object, ByVal proc As String) As Boolean
    Dim ln As Long
    On Error Resume Next
    ln = cm.ProcStartLine(proc, 0)
    HasProc = (Err.Number = 0 And ln > 0)
    On Error GoTo 0
End Function

' One packed argument (see the top) -> its value.
Private Function Unpack(ByVal s As String) As Variant
    Select Case Left$(s, 1)
        Case "s": Unpack = Mid$(s, 2)
        Case "n": Unpack = Val(Mid$(s, 2))
        Case "b": Unpack = (Mid$(s, 2) = "1")
        Case Else: Err.Raise vbObjectError + 7, "ScriptRun", "Bad packed argument '" & s & "'."
    End Select
End Function

' Token -> object (or the value itself, untouched). Into Dst rather than a
' return value, because Let on an object would take its default property.
Private Sub Resolve(ByVal v As Variant, ByRef Dst As Variant)
    Dim s As String, kind As String, arg As String, k As Long, ws As Worksheet, lo As ListObject

    If VarType(v) <> vbString Then
        Dst = v
        Exit Sub
    End If
    s = v
    If Left$(s, Len(TOK)) <> TOK Then
        Dst = s
        Exit Sub
    End If

    s = Mid$(s, Len(TOK) + 1)
    k = InStr(s, ":")
    If k = 0 Then
        kind = s
    Else
        kind = Left$(s, k - 1)
        arg = Mid$(s, k + 1)
    End If

    Select Case kind
        Case "wb"
            Set Dst = ThisWorkbook
        Case "sheet"
            Set Dst = ThisWorkbook.Worksheets(arg)
        Case "table"
            For Each ws In ThisWorkbook.Worksheets
                For Each lo In ws.ListObjects
                    If StrComp(lo.Name, arg, vbTextCompare) = 0 Then
                        Set Dst = lo
                        Exit Sub
                    End If
                Next lo
            Next ws
            Err.Raise vbObjectError + 2, "ScriptRun", "No table named '" & arg & "'."
        Case "range", "spill"
            k = InStrRev(arg, "!")
            If k = 0 Then Err.Raise vbObjectError + 3, "ScriptRun", "'" & arg & "' needs Sheet!Address."
            If kind = "range" Then
                Set Dst = ThisWorkbook.Worksheets(Left$(arg, k - 1)).Range(Mid$(arg, k + 1))
            Else
                Set Dst = ThisWorkbook.Worksheets(Left$(arg, k - 1)).Range(Mid$(arg, k + 1)).SpillingToRange
            End If
        Case "shape"
            k = InStrRev(arg, "!")
            If k = 0 Then Err.Raise vbObjectError + 3, "ScriptRun", "'" & arg & "' needs Sheet!ShapeName."
            Set Dst = ThisWorkbook.Worksheets(Left$(arg, k - 1)).Shapes(Mid$(arg, k + 1)).OLEFormat.Object
        Case "h"
            If mHandles Is Nothing Then Err.Raise vbObjectError + 4, "ScriptRun", "No handle " & arg & " (none kept)."
            Keep Dst, mHandles(CLng(arg))
        Case Else
            Err.Raise vbObjectError + 5, "ScriptRun", "Unknown token '" & TOK & kind & "'."
    End Select
End Sub

Private Function AddHandle(ByVal v As Variant) As String
    If mHandles Is Nothing Then Set mHandles = New Collection
    mHandles.Add v
    AddHandle = TOK & "h:" & mHandles.Count
End Function

' Dst = Src, by Set for an object. Src is a ByVal Variant so an object
' passed in arrives as itself, not as its default property.
Private Sub Keep(ByRef Dst As Variant, ByVal Src As Variant)
    If IsObject(Src) Then
        Set Dst = Src
    Else
        Dst = Src
    End If
End Sub
