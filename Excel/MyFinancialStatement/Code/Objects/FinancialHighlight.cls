VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Global Variable
Private myInfo As New SheetInfo
Private LookupSheet As Range
Private LookupRange As String

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-19
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    myInfo.Init FinancialHighlight.Cells, 1, "B"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GenerateData
' Author    : guhungry
' Date      : 2010-07-19
' Purpose   : Generate Financial Highlights from All Financial Statement
'---------------------------------------------------------------------------------------
'
Public Sub GenerateData()
    Dim Target As Range, Source As Range, Found As Range
    Dim balanceInfo As New SheetInfo
    Dim i As Long
    Dim LastCol As String
    
    Initialize
    balanceInfo.Init BalanceSheet.Cells
    'Delete Old Data
    If myInfo.LastCol > 2 Then
        myInfo.GetRange("C1").Delete xlShiftToLeft
    End If
    
    'No Data Do Nothing
    If balanceInfo.LastCol = 1 Then
        Exit Sub
    End If
    
    'Fill Financial Data
    FinancialHighlight.Cells(1, 3).Value = "=" & Utils.RangeAddress(BalanceSheet.Cells(1, 2))
    InitializeFindData BalanceSheet.Cells
    'Set Source = BalanceSheet.Cells
    For i = 4 To myInfo.LastRow
        Set Target = FinancialHighlight.Cells(i, 2)
        If Target.Value = "Income Statement" Then
            InitializeFindData IncomeStatement.Cells
            'Set Source = IncomeStatement.Cells
        ElseIf Target.Value = "Cash Flow Statement" Then
            InitializeFindData CashFlow.Cells
            'Set Source = CashFlow.Cells
        Else
            'Find Data
            Set Found = FindData(Target.Value)
            If Not (Found Is Nothing) Then
                Target(1, 2).Value = "=" & Utils.RangeAddress(Found(1, 2))
            End If
        End If
    Next i
    
    'Styling
    FinancialHighlight.Range("B1:B" & myInfo.LastRow).Copy
    FinancialHighlight.Range("C1:C" & myInfo.LastRow).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    'Expand Data
    LastCol = Utils.ColumnName(balanceInfo.CurCell(1, balanceInfo.LastCol + 1))
    If balanceInfo.LastCol > 2 Then
        FinancialHighlight.Range("C1:C" & myInfo.LastRow).Copy
        FinancialHighlight.Range("D1:" & LastCol & myInfo.LastRow).PasteSpecial xlAll
        Application.CutCopyMode = False
    End If
    myInfo.RestoreActiveCell
 End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitializeFindData
' Author    : guhungry
' Date      : 2010-07-19
' Purpose   : Initialize for Find Data
' @param Source    the data source
'---------------------------------------------------------------------------------------
'
Public Sub InitializeFindData(Source As Range)
    Dim info As New SheetInfo
    
    Set LookupSheet = Source
    info.Init LookupSheet
    LookupRange = "A1:A" & info.LastRow
End Sub

'---------------------------------------------------------------------------------------
' Function : FindData
' Author    : guhungry
' Date      : 2010-07-19
' Purpose   : Find Financial Data from source
' @param Value  the item value
' @return       matched cell
'---------------------------------------------------------------------------------------
'
Public Function FindData(Value As String) As Range
    Dim Found As Range
    
    If Value = "" Then
        Set FindData = Nothing
    Else
        Set Found = LookupSheet.Range(LookupRange)
        Set Found = Found.Find(What:=Value, After:=Found(1, 1), LookIn:=xlFormulas, LookAt _
            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        Set FindData = Found
    End If
End Function
