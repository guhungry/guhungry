VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Global Variable
Private myInfo As New SheetInfo

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-19
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    myInfo.Init FinancialRatio.Cells, 1, "B"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GenerateData
' Author    : guhungry
' Date      : 2010-07-19
' Purpose   : Generate Financial Ratio from Financial Highlights
'---------------------------------------------------------------------------------------
'
Public Sub GenerateData()
    Dim highlightInfo As New SheetInfo
    Dim LastCol As String
    
    Initialize
    highlightInfo.Init FinancialHighlight.Cells, 1, "B"
    'Delete Old Data
    If myInfo.LastCol > 2 Then
        myInfo.GetRange("C1").Delete xlShiftToLeft
    End If
    
    'No Data Do Nothing
    If highlightInfo.LastCol = 2 Then
        Exit Sub
    End If
    
    'Expand Data
    LastCol = Utils.ColumnName(highlightInfo.CurCell(1, highlightInfo.LastCol))
    FinancialRatio.Range("A1:A" & myInfo.LastRow).Copy
    FinancialRatio.Range("C1:" & LastCol & myInfo.LastRow).PasteSpecial xlAll
    Application.CutCopyMode = False
    myInfo.RestoreActiveCell
End Sub
