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

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' One printer, as loaded from the Printers sheet. A class rather than a Public
' Type for the same reason as clsStock.

Public PrinterID As String
Public Model As String
Public Consumable As String
Public RatePerM2 As Double
' Multi-pass costing (0.11.0). A blank or missing Colour mode reads as single
' pass and a blank Template cost as 0 - read-time defaults, nothing is written
' to older rows (design decision 34).
Public MultiPass As Boolean         ' Colour mode = "multi-pass"
Public TemplateCost As Double       ' start-up cost of one colour pass (e.g. a RISO master)
Public MaxRollWidthMM As Double     ' 0 = not roll-capable
Public MaxSheetSize As String       ' tblStandardSizes[Size name]; "" = not sheet-capable
Public MaxSheetWidthMM As Double    ' resolved from MaxSheetSize at load time
Public MaxSheetHeightMM As Double   ' "
Public Active As Boolean
Public Found As Boolean
