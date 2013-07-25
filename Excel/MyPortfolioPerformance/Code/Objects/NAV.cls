VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Global Variable
Private info As New SheetInfo

''''''''''''''''''''''
' Last Row property
''''''''''''''''''''''
Public Property Get LastRow() As Long
LastRow = info.LastRow
End Property

''''''''''''''''''''''
' Active Cell property
''''''''''''''''''''''
Public Property Get ActCell() As String
ActCell = info.ActCell
End Property

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author : guhungry
' Date : 2010-07-01
' Purpose : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
info.Init NAV.Cells, 2
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClearData
' Author : guhungry
' Date : 2012-07-01
' Purpose : Clear sheet data
'---------------------------------------------------------------------------------------
'
Public Sub ClearData()
    If LastRow > 2 Then
        NAV.Range("A3:I" & LastRow).Delete Shift:=xlUp
    End If
End Sub
