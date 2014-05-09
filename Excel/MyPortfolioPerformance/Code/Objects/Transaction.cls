VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transaction"
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

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author : guhungry
' Date : 2010-07-01
' Purpose : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    info.Init Transaction.Cells, 2, "A"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnNAVGenerate_Click
' Author : guhungry
' Date : 2012-12-31
' Purpose : Click Generate NAV Button
'---------------------------------------------------------------------------------------
'
Private Sub btnNAVGenerate_Click()
    Initialize
    SortTransaction
    
    GenerateNAVReport
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GenerateNAVReport
' Author : guhungry
' Date : 2012-12-31
' Purpose : Generate NAV Report from Transaction and Portfolio Data
'---------------------------------------------------------------------------------------
'
Private Sub GenerateNAVReport()
    Dim myActive As String
    Dim curDate As Date, prevDate As Date, endYear As Date
    Dim curStock As String
    Dim myComment As String
    Dim lYear As Integer, tYear As Integer
    Dim myCell As Range, endCell As Range, tran As Range, target As Range
    Dim LastRow As Long, EndRow As Long, StartRow As Long, NumStock As Long
    
    'Initialize Variables
    Application.ScreenUpdating = False
    NAV.EnableCalculation = False
    Initialize
    NAV.Initialize
    NAV.ClearData
    NAV.Activate
    Set myCell = NAV.Cells(3, 1)
    LastRow = 2
    Do While LastRow <= Transaction.LastRow
        Set tran = Transaction.Range("B" & LastRow)
        curDate = tran.Cells(1, 1).Value
        
        'Find EndRow
        EndRow = LastRow + 1
        Set endCell = Transaction.Range("C" & EndRow)
        Do While EndRow <= Transaction.LastRow
            Set endCell = Transaction.Range("C" & EndRow)
            If endCell.Cells(1, 0).Value <> curDate Then
                Exit Do
            End If
            EndRow = EndRow + 1
        Loop
        EndRow = EndRow - 1
        
        'Before Transaction Date
        prevDate = Evaluate("=MAX(IF(Data!A:A<" & Int(CDbl(curDate)) & ",Data!A:A))")
        If LastRow <> 2 And myCell.Cells(0, 1).Value <> prevDate Then
            myCell.Cells(1, 1).Value = prevDate
            
            Set myCell = myCell.Cells(2, 1)
        End If
        
        'Transaction Date
        myCell.Cells(1, 1).Value = curDate
        
        Set myCell = myCell.Cells(2, 1)
        
        'End Year Date
        If IsDate(tran.Cells(2, 1).Value) Then
            lYear = Year(Evaluate("=MIN(IF(Transaction!B:B>" & Int(CDbl(curDate)) & ",Transaction!B:B))"))
        Else
            lYear = Year(Date)
        End If
        If Year(curDate) <> lYear Then
            For tYear = Year(curDate) To lYear - 1
                endYear = Evaluate("=MAX(IF(Data!A:A<" & Int(CDbl(DateSerial(tYear, 12, 31))) & ",Data!A:A))")
                If endYear <> myCell.Cells(0, 1).Value Then
                    myCell.Cells(1, 1).Value = endYear
                    
                    Set myCell = myCell.Cells(2, 1)
                End If
            Next tYear
        End If
        
        LastRow = EndRow + 1
    Loop
    
    'Today
    prevDate = Evaluate("=MAX(IF(Data!A:A<=" & Int(CDbl(Date)) & ",Data!A:A))")
    If prevDate <> myCell.Cells(0, 1).Value Then
        myCell.Cells(1, 1).Value = prevDate
    End If
    
    myActive = NAV.ActCell
    NAVPasteFormula
    NAV.Range(myActive).Select
    Transaction.Activate
    NAV.EnableCalculation = True
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : NAVPasteFormula
' Author : guhungry
' Date : 2014-03-10
' Purpose : Copy formula template to NAV
'---------------------------------------------------------------------------------------
'
Private Sub NAVPasteFormula()
    Dim source As Range, target As Range
    
    NAV.Initialize
    If NAV.LastRow > 2 Then
        Application.ScreenUpdating = False
        Set source = NAV.Range("B2:I2")
        'Set target = NAV.Range("B3:I4")
        Set target = NAV.Range("B3:I" & NAV.LastRow)
        
        source.Copy
        target.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SortTransaction
' Author : guhungry
' Date : 2012-12-31
' Purpose : Sort Transaction Data By Date
'---------------------------------------------------------------------------------------
'
Private Sub SortTransaction()
    Dim target As Range
    
    Application.ScreenUpdating = False
    If info.LastRow > 10 Then
    Set target = Transaction.Range("A2:D" & info.LastRow)
    
    'Sort
    Transaction.Sort.SortFields.Clear
    Transaction.Sort.SortFields.Add Key:=Range("B2:B" & info.LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Transaction.Sort
        .SetRange target
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    info.RestoreActiveCell
    End If
    Application.ScreenUpdating = True
End Sub

