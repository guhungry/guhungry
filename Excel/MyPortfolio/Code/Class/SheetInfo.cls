VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Class : SheetInfo
' Author    : guhungry
' Date      : 2010-07-08
' Purpose   : Keep the sheet info (Last Row, Last Column, Current Row)
' Example   :   SheetInfo info As New SheetInfo
'               info.Init ActiveSheet.Cells, 1, "B"
'---------------------------------------------------------------------------------------
'
Private pSheet As Range
Private pLastRow As Long
Private pLastCol As Long
Private pCurRow As Long
Private pCurCell As Range
Private pActiveCell As String

''''''''''''''''''''''
' Last Row property
''''''''''''''''''''''
Public Property Get LastRow() As Long
    LastRow = pLastRow
End Property
Private Property Let LastRow(Value As Long)
    pLastRow = Value
End Property

''''''''''''''''''''''
' Last Column property
''''''''''''''''''''''
Public Property Get LastCol() As Long
    LastCol = pLastCol
End Property
Private Property Let LastCol(Value As Long)
    pLastCol = Value
End Property

''''''''''''''''''''''
' Current Row property
''''''''''''''''''''''
Public Property Get CurRow() As Long
    CurRow = pCurRow
End Property
Public Property Let CurRow(Value As Long)
    pCurRow = Value
    If pCurRow <= LastRow And LastRow > 1 Then
        Set pCurCell = pSheet.Range("A" & pCurRow)
    Else
        Set pCurCell = Nothing
    End If
End Property

''''''''''''''''''''''
' Current Cell property
''''''''''''''''''''''
Public Property Get CurCell() As Range
    Set CurCell = pCurCell
End Property

''''''''''''''''''''''
' Active Cell property
''''''''''''''''''''''
Public Property Get ActCell() As String
    ActCell = pActiveCell
End Property

'---------------------------------------------------------------------------------------
' Procedure : Init
' Author    : guhungry
' Date      : 2010-07-13
' Purpose   : Initialize Sheet Info
'---------------------------------------------------------------------------------------
'
Public Sub Init(SheetName As Range, Optional Row As Long = 1, Optional Column As String = "A")
    Set pSheet = SheetName
    LastCol = Utils.GetLastColumn(pSheet, Row)
    LastRow = Utils.GetLastRow(pSheet, Column)
    CurRow = 1

    'Get Active Cell
    pSheet.Parent.Activate
    pActiveCell = ActiveCell.Address
End Sub

'---------------------------------------------------------------------------------------
' Function  : GetRange
' Author    : guhungry
' Date      : 2010-07-14
' Purpose   : Get Sheet Range
' @param StartCell  the first cell of range
' @return           returns sheet range
'---------------------------------------------------------------------------------------
'
Public Function GetRange(Optional StartCell As String = "A1") As Range
    Set GetRange = pSheet.Range(StartCell & ":" & Utils.ColumnName(pSheet(1, LastCol)) & LastRow)
End Function

'---------------------------------------------------------------------------------------
' Procedure : RestoreActiveCell
' Author    : guhungry
' Date      : 2010-08-02
' Purpose   : Restore ActiveCell in Sheet
'---------------------------------------------------------------------------------------
'
Public Sub RestoreActiveCell()
    pSheet.Parent.Activate
    pSheet.Range(pActiveCell).Select
End Sub