VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRestoreRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' One print room's share of a restore (modBackup): what the backup says about
' the room, read up front, plus - once modBackup has found or created the room
' in this workbook - the sheet its job records go into.

Public Code As String            ' LOC_Code; empty when the source does not say
Public RoomName As String        ' LOC_Name
Public Dept As String            ' LOC_Dept
Public JobRows As Collection     ' clsDict per job record, keyed by header (modImport's shape)
Public Settings As clsDict       ' the room's LOC_* settings by name; Nothing when the source has none (a CSV backup)
Public Sheet As Worksheet        ' the room in THIS workbook; Nothing until found/created
Public Created As Boolean        ' True when the restore had to create it
