VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' A keyed store built on Collection.
'
' Excel for Mac has no ActiveX and no CreateObject, so Scripting.Dictionary is
' unavailable there. Every lookup table in this project goes through this class
' instead, which behaves the same on both platforms.
'
' Keys are case-insensitive, matching how users actually type printer and stock
' names.

Private mItems As Collection
Private mKeys As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
    Set mKeys = New Collection
End Sub

Private Function Norm(ByVal Key As String) As String
    Norm = LCase$(Trim$(Key))
End Function

Public Sub Add(ByVal Key As String, ByVal Value As Variant)
    Dim k As String
    k = Norm(Key)
    If Exists(k) Then Remove k
    mItems.Add Value, k
    mKeys.Add Key, k
End Sub

Public Function Exists(ByVal Key As String) As Boolean
    ' Referencing the item is the whole test: a missing key raises and lands on
    ' NotFound. IsObject is used rather than assigning to a Variant, because a
    ' plain "v = mItems(...)" is a Let assignment, and letting an object into a
    ' Variant without Set raises error 438. That silently made every lookup for
    ' an object value report False.
    Dim probe As Boolean
    On Error GoTo NotFound
    probe = IsObject(mItems(Norm(Key)))
    Exists = True
    Exit Function
NotFound:
    Exists = False
End Function

Public Function Item(ByVal Key As String) As Variant
    ' Same trap as Exists: an object value has to be handed back with Set.
    On Error GoTo Missing
    If IsObject(mItems(Norm(Key))) Then
        Set Item = mItems(Norm(Key))
    Else
        Item = mItems(Norm(Key))
    End If
    Exit Function
Missing:
    Item = Empty
End Function

' Objects have to be fetched with Set, so they get their own accessor rather
' than relying on a Variant return doing the right thing in every position.
Public Function Obj(ByVal Key As String) As Object
    On Error GoTo Missing
    Set Obj = mItems(Norm(Key))
    Exit Function
Missing:
    Set Obj = Nothing
End Function

Public Sub Remove(ByVal Key As String)
    On Error Resume Next
    mItems.Remove Norm(Key)
    mKeys.Remove Norm(Key)
    On Error GoTo 0
End Sub

Public Property Get Count() As Long
    Count = mItems.Count
End Property

' Original-case keys, in insertion order. Returns an empty array when the
' store is empty, so callers can loop without a special case.
Public Function Keys() As Variant
    Dim a() As String, i As Long
    If mKeys.Count = 0 Then
        Keys = Array()
        Exit Function
    End If
    ReDim a(0 To mKeys.Count - 1)
    For i = 1 To mKeys.Count
        a(i - 1) = mKeys(i)
    Next i
    Keys = a
End Function
