VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewData"
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
' Date      : 2010-07-15
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    myInfo.Init NewData.Cells, 2
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GenerateFinancialData
' Author    : guhungry
' Date      : 2010-07-19
' Purpose   : Generate Financial Highlights & Financial Ratios
'---------------------------------------------------------------------------------------
'
Public Sub GenerateFinancialData()
    Application.ScreenUpdating = False
    FinancialHighlight.GenerateData
    FinancialRatio.GenerateData
    NewData.Activate
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MergeBalanceSheet
' Author    : guhungry
' Date      : 2010-07-15
' Purpose   : Merge New Data with Balance Sheet
'---------------------------------------------------------------------------------------
'
Public Sub MergeBalanceSheet()
    MergeStatement BalanceSheet.Cells, "A"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MergeIncomeStatement
' Author    : guhungry
' Date      : 2010-07-15
' Purpose   : Merge New Data with Income Statement
'---------------------------------------------------------------------------------------
'
Public Sub MergeIncomeStatement()
    MergeStatement IncomeStatement.Cells, "B"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MergeCashFlowStatement
' Author    : guhungry
' Date      : 2010-07-15
' Purpose   : Merge New Data with Cash Flow Statement
'---------------------------------------------------------------------------------------
'
Public Sub MergeCashFlowStatement()
    MergeStatement CashFlow.Cells, "C"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MergeStatement
' Author    : guhungry
' Date      : 2010-07-15
' Purpose   : Base merge function (Clean up column, merge stament)
'---------------------------------------------------------------------------------------
'
Private Sub MergeStatement(Source As Range, LookupColumn As String)
    Initialize
    Dim Cell As Range
    Dim col As String
    Dim sourceInfo As New SheetInfo
    Dim myActive As String
    Dim myCell As Range
    Dim LastColumn As Long
    
    ' 1) Filter Header Column (Year) - Example: Yearly/2012
    Set myCell = NewData.Cells(1, 2)
    Do While myCell.Column <= myInfo.LastCol
        myCell.Value = TextUtils.GetRegExp("[0-9]+$", myCell.Value)
    
        Set myCell = myCell.Cells(1, 2)
    Loop
    
    ' Detect if Data from Set Smart
    ' 1) Delete second row - Example: (31/12/2012)
    ' 2) Copy not empty cell in column B to column A
    ' 3) Delete column B
    If IsEmpty(NewData.Cells(3, 1).Value) Then
        ' 1) Delete second row - Example: (31/12/2012)
        NewData.Range("2:2").Delete
        
        ' 2) Copy not empty cell in column B to column A
        Set myCell = NewData.Cells(3, 2)
        Do While myCell.Row <= myInfo.LastRow - 1
            If Not IsEmpty(myCell.Value) Then
                myCell(1, 0).Value = myCell.Value
            End If
            Set myCell = myCell.Cells(2, 1)
        Loop
        
        ' 3) Delete column B
        NewData.Range("B:B").Delete
    End If


    'Clean up unneeded columns
    Application.ScreenUpdating = False
    Do
        Set Cell = FindUnneededColumn()
        If Cell Is Nothing Then
            Exit Do
        End If
        col = Utils.ColumnName(Cell)
        NewData.Range(col & ":" & col).Delete
        NewData.Cells(1, myInfo.LastCol + 1).EntireColumn.Insert
    Loop
    
    'Merge Financial Statement
    sourceInfo.Init Source.Cells
    myActive = sourceInfo.ActCell
    FStatement.MergeStatement Source.Cells, NewData.Cells, LookupColumn
    
    'Restore Active Cell
    Temp.Activate
    Temp.Range("A1").Select
    Source.Parent.Activate
    Source.Range(myActive).Select
    NewData.Activate
    NewData.Range(myInfo.ActCell).Select
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Function  : FindUnneededColumn
' Author    : guhungry
' Date      : 2010-07-15
' Purpose   : Return cell of matched Item
' @return       unneeded cell
'---------------------------------------------------------------------------------------
'
Private Function FindUnneededColumn() As Range
    Dim Found As Range
    Set Found = NewData.Range("B2:" & Utils.ColumnName(NewData.Cells(1, myInfo.LastCol)) & "2")
    Set Found = Found.Find(What:="%", After:=Found(1, 1), LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Set FindUnneededColumn = Found
End Function
