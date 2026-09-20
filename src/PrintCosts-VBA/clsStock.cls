VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One paper stock, as loaded from the Papers sheet.
'
' This is a class rather than a Public Type because VBA cannot put a
' user-defined type into a Variant or a Collection when the type is declared in
' a standard module - "Only user-defined types defined in public object modules
' can be coerced to or from a variant". The catalogue keeps these in a clsDict,
' which is a Collection underneath, so they have to be objects.

Public StockID As String
Public Description As String
Public PaperType As String
Public Family As String
Public Measure As String            ' "Sheet" or "Roll"
Public WidthMM As Double
Public HeightMM As Double
Public Cost As Double               ' per sheet, or per metre
Public Active As Boolean
Public Found As Boolean
