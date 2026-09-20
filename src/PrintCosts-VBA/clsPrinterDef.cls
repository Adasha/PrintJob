VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrinterDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One printer, as loaded from the Printers sheet. A class rather than a Public
' Type for the same reason as clsStock.

Public PrinterID As String
Public Model As String
Public Consumable As String
Public RatePerM2 As Double
Public Families As String           ' delimited with LIST_SEP
Public Active As Boolean
Public Found As Boolean
