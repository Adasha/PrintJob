VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' One colour, as loaded from the Consumables sheet. A class for the same
' reason as clsStock and clsPrinterDef.

Public ColourID As String
Public ConsumableType As String     ' tblConsumables[Consumable type], e.g. "Risograph"
Public Colour As String
Public RatePerM2 As Double
Public Active As Boolean
Public Found As Boolean
