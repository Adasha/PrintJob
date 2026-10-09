VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
' If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.

' One department, as loaded from the Departments sheet. A class for the same
' reason clsStock is: the catalogue keeps these in a clsDict (a Collection
' underneath), which cannot hold a Public Type declared in a standard module.
' The same object is stored under the department's Name and under each of its
' Aliases, so one lookup resolves either spelling.

Public DeptID As String
Public Name As String
Public Free As Boolean          ' Free = Yes: a job for this department starts with both disregards set to Yes
Public Active As Boolean
Public Found As Boolean
