VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Data"
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
    LastRow = IFF(info.LastRow = 1, 2, info.LastRow)
End Property

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author : guhungry
' Date : 2013-07-29
' Purpose : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    info.Init Data.Cells, 1, "A"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnGraphGenerate_Click
' Author : guhungry
' Date : 2013-07-29
' Purpose : Generate Portfolio Performance Chart (Change Data Source)
'---------------------------------------------------------------------------------------
'
Sub btnGraphGenerate_Click()
    Dim mchart As ChartObject
    Initialize

    Set mchart = Data.ChartObjects("ChartPerformance")
    mchart.chart.SetSourceData Source:=Range("Data!$A$1:$F$" & LastRow)
End Sub


