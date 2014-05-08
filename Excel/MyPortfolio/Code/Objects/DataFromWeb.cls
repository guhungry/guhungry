VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFromWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Procedure : btnResetWebData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Clear DataFromWeb and regenerate DataFromWeb
' Change Log
'  yyyy-mm-dd  - Fix Fucked Up Web Data - Cause of Listed or Delisted stock(s).
' 2013-05-25 - Add Warrant Table
'---------------------------------------------------------------------------------------
'
Sub btnResetWebData()
    'Define Variables
    Dim LastRow As Integer
    LastRow = Utils.GetLastRow(DataFromWeb.Cells) + 2

    'Clear Data
    Application.ScreenUpdating = False
    Range("A1:L" & LastRow).Delete Shift:=xlUp
    LastRow = 1
    'Delete all connections
    For Each cn In ThisWorkbook.Connections
        cn.Delete
    Next cn
    'Delete data from web query
    While DataFromWeb.QueryTables.Count
        DataFromWeb.QueryTables(1).Delete
    Wend
    'SET Main
    SetDataFromWeb "SetMain", GetNameRefersTo(Setting.Names("URL_SETMAIN")), GetNameRefersTo(Setting.Names("TAB_SETMAIN")), "A" & LastRow
    LastRow = Utils.GetLastRow(DataFromWeb.Cells) + 2
    'SET Foreign
    SetDataFromWeb "SetForeign", GetNameRefersTo(Setting.Names("URL_SETFOREIGN")), GetNameRefersTo(Setting.Names("TAB_SETFOREIGN")), "A" & LastRow
    LastRow = Utils.GetLastRow(DataFromWeb.Cells) + 2
    'SET Warrant
    SetDataFromWeb "SetWarrant", GetNameRefersTo(Setting.Names("URL_SETWARRANT")), GetNameRefersTo(Setting.Names("TAB_SETWARRANT")), "A" & LastRow
    LastRow = Utils.GetLastRow(DataFromWeb.Cells) + 2
    'MAI Main
    SetDataFromWeb "MaiMain", GetNameRefersTo(Setting.Names("URL_MAIMAIN")), GetNameRefersTo(Setting.Names("TAB_MAIMAIN")), "A" & LastRow
    LastRow = Utils.GetLastRow(DataFromWeb.Cells) + 2
    'MAI Foreign
    SetDataFromWeb "MaiForeign", GetNameRefersTo(Setting.Names("URL_MAIFOREIGN")), GetNameRefersTo(Setting.Names("TAB_MAIFOREIGN")), "A" & LastRow
    LastRow = Utils.GetLastRow(DataFromWeb.Cells) + 2
    'MAI Warrant
    SetDataFromWeb "MaiWarrant", GetNameRefersTo(Setting.Names("URL_MAIWARRANT")), GetNameRefersTo(Setting.Names("TAB_MAIWARRANT")), "A" & LastRow
    LastRow = Utils.GetLastRow(DataFromWeb.Cells) + 2
    
    Transaction.btnTransactionSortByTicker
    DataFromWeb.Activate
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnRefreshFromWeb
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Update Data From Web Now!!!!
'---------------------------------------------------------------------------------------
'
Sub btnRefreshFromWeb()
    ActiveWorkbook.RefreshAll
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnUpdateStockPrice
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Update Web Data to StockPrice
'---------------------------------------------------------------------------------------
'
Sub btnUpdateStockPrice()
    'Define Variables
    Dim NumStockPrice As Long
    Dim list As Range
    Dim price As Double
    Dim today As Date

    'Default Value
    Stock.Initialize
    StockPrice.Range("A1").Value = DataFromWeb.Range("A2").Value
    Set list = StockPrice.Range("A6:A" & Application.WorksheetFunction.Max(Utils.GetLastRow(StockPrice.Cells), 6))
    today = GetUpdateDate()

    ' Update Stock Price
    For Each S In list
        If DateDiff("d", S.Cells(1, 2).Value, today) > 0 Then
            price = Stock.GetStockValue(S.Value)
            If price > 0 Then
                S.Cells(1, 2).Value = today
                S.Cells(1, 3).Value = price
            End If
        End If
    Next
    
    'Insert Unknown StockPrice
    Set list = DataFromWeb.Range("A28:A" & Stock.LastDataFromWeb)
    If (Stock.LastStockPrice > 5) Then
        For Each S In list
            If IsNumber(S.Cells(1, 6).Value) And Not (Stock.IsStockPriceExist(S.Value)) Then
                price = S.Cells(1, 6).Value
                Stock.InsertStockPrice S.Value, today, price
            End If
        Next
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetDataFromWeb
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Set up DataFromWeb connection to given target cell
' @param Name   the connection name
' @param Table  the data table
' @param target the target cell
'---------------------------------------------------------------------------------------
'
Private Sub SetDataFromWeb(Name As String, URL As String, Table As String, target As String)
    With DataFromWeb.QueryTables.Add(Connection:="URL;" & URL, Destination:=DataFromWeb.Range(target))
        .Name = Name
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingAll
        .WebTables = Table
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub
