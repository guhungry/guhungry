VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Portfolio"
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
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    info.Init Portfolio.Cells, 23, "B"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PortfolioMake
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Generate portfolio, cost, current value and graph.
'---------------------------------------------------------------------------------------
'
Sub PortfolioMake()
    Dim source As Range, target As Range
    Dim oldList() As Long, curList() As Long
    Dim numOldStock As Long, numNewStock As Long
    Dim HasNewStock As Boolean
    Dim rowOldHeader As Long
    Dim uiLang As String

    'Initialize Variables
    ReDim oldList(0 To 0)
    ReDim curList(0 To 0)
    Stock.Initialize
    Report.Initialize
    uiLang = GetNameRefersTo(Setting.Names("uiLang"))

    'Parse Stock Data
    ParseTransaction curList(), oldList()
    HasNewStock = (UBound(curList) > 0)

    'Print Current Stocks
    PrintStock 27, curList(), True

    'Print Old Stocks
    rowOldHeader = 27 + UBound(curList)
    PrintStock rowOldHeader + 1, oldList(), False

    'Copy Price
    Initialize
    If LastRow > 26 Then
        Set source = Portfolio.Range("I27:I" & LastRow)
        Set target = Portfolio.Range("J27:O" & LastRow)
        source.Copy
        target.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If

    'Add Header
    numNewStock = Application.WorksheetFunction.Max(UBound(curList), 1)
    numOldStock = Application.WorksheetFunction.Max(UBound(oldList), 1)
    'Old Stock Header
    Portfolio.Cells(rowOldHeader, 2).Value = GetNameRefersTo(Setting.Names("txtPastStock_" & uiLang))
    Portfolio.Cells(rowOldHeader, 3).Value = "=SUM(R[1]C:R[" & numOldStock & "]C)"
    'Current Stock Header
    Portfolio.Cells(26, 2).Value = GetNameRefersTo(Setting.Names("txtCurrentStock_" & uiLang))
    If HasNewStock Then
        Portfolio.Cells(26, 3).Value = "=SUM(R[1]C:R[" & numNewStock & "]C)"
        Portfolio.Cells(26, 4).Value = "=SUM(R[1]C:R[" & numNewStock & "]C)"
        Portfolio.Cells(26, 5).Value = "=SUM(R[1]C:R[" & numNewStock & "]C)"
        Portfolio.Cells(26, 6).Value = "=IF(RC[-3]<=0,R[" & (1 + numNewStock) & "]C[-3]*0.01,RC[-1]/RC[-3])"
    Else
        Portfolio.Cells(26, 3).Value = 0
        Portfolio.Cells(26, 4).Value = 0
        Portfolio.Cells(26, 5).Value = 0
        Portfolio.Cells(26, 6).Value = 0
    End If
    'Cash
    Portfolio.Cells(25, 2).Value = GetNameRefersTo(Setting.Names("txtCash_" & uiLang))
    Portfolio.Cells(25, 3).FormulaR1C1 = "=R[-1]C - R[1]C - R[" & (2 + numNewStock) & "]C"
    Portfolio.Cells(25, 4).FormulaR1C1 = "=RC[-1]"
    'Total Asset
    Portfolio.Cells(24, 2).Value = GetNameRefersTo(Setting.Names("txtNetValue_" & uiLang))
    Portfolio.Cells(24, 3).Value = Stock.FundSource
    Portfolio.Cells(24, 4).FormulaR1C1 = "=R[1]C+R[2]C"
    Portfolio.Cells(24, 5).FormulaR1C1 = "=RC[-1]-RC[-2]"
    Portfolio.Cells(24, 6).FormulaR1C1 = "=IF(RC[-3]<=0,RC[-1]*0.01,RC[-1]/RC[-3])"
    Portfolio.Cells(24, 15).Value = Stock.GetLastUpdate()

    'Charts
    Portfolio.ChartObjects("CurrentCost").Activate
    ActiveChart.SetSourceData source:=Portfolio.Range("B25:C25" & Utils.IFF(HasNewStock, ",B27:C" & (26 + numNewStock), "")), PlotBy:=xlColumns
    Portfolio.ChartObjects("CurrentValue").Activate
    ActiveChart.SetSourceData source:=Portfolio.Range("B25:D25" & Utils.IFF(HasNewStock, ",B27:D" & (26 + numNewStock), "")), PlotBy:=xlColumns
    ActiveChart.SeriesCollection(2).Delete
    ActiveChart.SeriesCollection(1).Values = "=Portfolio!$D$25" & Utils.IFF(HasNewStock, ",Portfolio!$D$27:$D$" & (26 + numNewStock), "")

    'Add Style
    Initialize
    PortfolioStyleAdd
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintStock
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Print stock data in Portfolio
'---------------------------------------------------------------------------------------
'
Public Sub PrintStock(FirstRow As Long, ListStock() As Long, IsNewStock As Boolean)
    Dim i As Long
    Dim myCell As Range, myStock As Range
    Dim myComment As String

    Set myCell = Portfolio.Range("B" & FirstRow)
    For i = 1 To UBound(ListStock)
        Set myStock = Report.Range("C" & ListStock(i))
        If IsNewStock Then
            'Cost Data
            '# of Share
            myCell.Cells(1, 6).Value = myStock.Cells(1, 3).Value
            myCell.Cells(1, 7).Value = myStock.Cells(1, 4).Value
            myCell.Cells(1, 8).Formula = myStock.Cells(1, 5).Formula
            myCell.Cells(1, 8).AddComment (myStock.Cells(1, 5).Comment.Text)
            myCell.Cells(1, 2).Value = "=RC[4]*RC[5]"
            myCell.Cells(1, 3).Value = "=RC[3]*RC[5]"
            myCell.Cells(1, 4).Value = "=RC[-1]-RC[-2]"
            myCell.Cells(1, 5).Value = "=RC[-1]/RC[-3]"
        Else
            'Cost Data
            myCell.Cells(1, 2).Value = -myStock.Cells(1, 2).Value

            'Latest Stock Price
            myCell.Cells(1, 8).Value = Stock.GetLatestStockPrice(myStock.Value, myComment)
            myCell.Cells(1, 8).AddComment myComment
        End If

        'Stock Name
        myCell.Value = myStock.Value
        myCell.Cells(1, 8).Comment.Shape.TextFrame.AutoSize = True

        Set myCell = myCell.Cells(2, 1)
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PortfolioClearData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Clear portfolio data
'---------------------------------------------------------------------------------------
'
Public Sub PortfolioClearData()
    'Clear Graph
    Portfolio.ChartObjects("CurrentCost").Activate
    ActiveChart.SetSourceData source:=Portfolio.Range("B23:C23")
    Portfolio.ChartObjects("CurrentValue").Activate
    ActiveChart.SetSourceData source:=Portfolio.Range("B23:C23")

    'Clear Data
    If LastRow > 23 Then
        Portfolio.Range("B24:O" & LastRow).Delete Shift:=xlUp
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InsertRow
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Insert blank row for portfolio
'---------------------------------------------------------------------------------------
'
Private Sub InsertRow(RowNum As Long)
    Portfolio.Range("A" & RowNum).EntireRow.Insert xlDown, xlFormatFromRightOrBelow
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportStyleAdd
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Copy styling template to portfolio data
'---------------------------------------------------------------------------------------
'
Private Sub PortfolioStyleAdd()
    Dim source As Range, target As Range

    If LastRow > 23 Then
        Set source = Portfolio.Range("B22:O22")
        Set target = Portfolio.Range("B24:O" & LastRow)

        source.Copy
        target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParseTransaction
' Author    : guhungry
' Date      : 2010-07-06
' Purpose   : Get current stock list and old stock list
'---------------------------------------------------------------------------------------
'
Private Sub ParseTransaction(ByRef listNew() As Long, ByRef listOld() As Long)
    Dim iNew As Long, iOld As Long
    Dim Row As Long
    iNew = 1
    iOld = 1

    For Row = 6 To Report.LastRow
        If Not (IsEmpty(Report.Cells(Row, 3))) Then
            If Report.Cells(Row, 5).Value > 0 Then
                ReDim Preserve listNew(iNew)
                listNew(iNew) = Row
                iNew = iNew + 1
            Else
                ReDim Preserve listOld(iOld)
                listOld(iOld) = Row
                iOld = iOld + 1
            End If
        End If
    Next Row
End Sub
