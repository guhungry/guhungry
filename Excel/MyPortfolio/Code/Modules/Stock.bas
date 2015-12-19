Attribute VB_Name = "Stock"
'Global Variable
Public LastDataFromWeb As Long
Public LastStockPrice As Long
Public FundSource As String
Private IsDataFromWebNewer As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    LastDataFromWeb = Utils.GetLastRow(DataFromWeb.Cells)
    LastStockPrice = Utils.GetLastRow(StockPrice.Cells)
    IsDataFromWebNewer = (DateDiff("d", GetUpdateDate("STOCK"), GetUpdateDate("WEB")) >= 0)
    FundSource = "=" & Utils.RangeAddress(Fund.Range("F10"))
End Sub

'---------------------------------------------------------------------------------------
' Function  : FindStock
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return Range of matched Stock in DataFromWeb, used by GetStockValue and GetStockCell
'---------------------------------------------------------------------------------------
'
Private Function FindStock(StockName As String) As Range
    Dim Found As Range
    Set Found = DataFromWeb.Range("A27", "A" & LastDataFromWeb)
    Set Found = Found.Find(What:=StockName & " ", After:=Found.Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    Set FindStock = Found
End Function

'---------------------------------------------------------------------------------------
' Function  : GetStockValue
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return price of matched Stock in DataFromWeb
' @param StockName  the stock name
' @return           the price of stock
' Example           GetStockValue("AIT") => 29.5
'---------------------------------------------------------------------------------------
'
Public Function GetStockValue(StockName As String) As Double
    Dim Found As Range
    Set Found = FindStock(StockName)

    If Found Is Nothing Then
        GetStockValue = -1
    Else
        If Utils.IsNumber(Found.Cells(1, 6)) Then
            GetStockValue = Found.Cells(1, 6).Value
        Else
            GetStockValue = -1
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : GetStockCell
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return address of matched Stock in DataFromWeb
' @param StockName  the stock name
' @return           the address of stock price cell
' Example           GetStockCell("AIT") => 'DataFromWeb!F6'
'---------------------------------------------------------------------------------------
'
Public Function GetStockCell(StockName As String) As String
    Dim Found As Range
    Set Found = FindStock(StockName)

    If Found Is Nothing Then
        GetStockCell = ""
    Else
        If Utils.IsNumber(Found.Cells(1, 6)) Then
            GetStockCell = "=" & Utils.RangeAddress(Found.Cells(1, 6))
        Else
            GetStockCell = ""
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : GetUpdateDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return update date of DataFromWeb or StockPrice
' @param SheetName  the Sheet Name ('WEB', 'STOCKPRICE')
' @return           the address of stock price cell
' Example           GetUpdateDate("WEB") => '30 June 2010'
'---------------------------------------------------------------------------------------
'
Public Function GetUpdateDate(Optional SheetName As String = "WEB") As Date
    If SheetName = "WEB" Then
        GetUpdateDate = TextUtils.ExtractDate(DataFromWeb.Range(GetNameRefersTo(Setting.Names("DATE_WEB"))).Value)
    Else
        GetUpdateDate = TextUtils.ExtractDate(StockPrice.Range(GetNameRefersTo(Setting.Names("DATE_STOCKPRICE"))).Value)
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : GetLastUpdate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return latest update date from DataFromWeb or StockPrice
' @return           the address of stock price cell
' Example           GetLastUpdate() => 'Last Update 30 Jun 2010 16:59:45'
'---------------------------------------------------------------------------------------
'
Public Function GetLastUpdate() As String
    If IsDataFromWebNewer Then
        GetLastUpdate = DataFromWeb.Range(GetNameRefersTo(Setting.Names("DATE_WEB"))).Value
    Else
        GetLastUpdate = StockPrice.Range(GetNameRefersTo(Setting.Names("DATE_STOCKPRICE"))).Value
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : IsStockPriceExist
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Check weather stock exist in StockPrice
' @param StockName  the stock name
' @return           Is stock exists in StockPrice
' Example           IsStockPriceExist("AIT") => True
'---------------------------------------------------------------------------------------
'
Public Function IsStockPriceExist(StockName As String) As Boolean
    Dim Found As Range

    Set Found = StockPrice.Range("A5:A" & LastStockPrice)
    Set Found = Found.Find(What:=Trim(StockName), After:=Found.Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

    If Found Is Nothing Then
        IsStockPriceExist = False
    Else
        IsStockPriceExist = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : GetStockPriceCell
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return address of matched Stock in StockPrice
' @param StockName  the stock name
' @return           the address of stock price cell
' Example           GetStockCell("AIT") => 'StockPrice!C35'
'---------------------------------------------------------------------------------------
'
Public Function GetStockPriceCell(StockName As String) As String
    Dim Found As Range

    Set Found = StockPrice.Range("A5:A" & LastStockPrice)
    Set Found = Found.Find(What:=Trim(StockName), After:=Found.Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If Found Is Nothing Then
        GetStockPriceCell = ""
    Else
        If Utils.IsNumber(Found.Cells(1, 3)) Then
            GetStockPriceCell = "=" & Utils.RangeAddress(Found.Cells(1, 3))
        Else
            GetStockPriceCell = ""
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : InsertStockPrice
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Insert Stock into StockPrice
' @param StockName  the stock name
' @param StockDate  the update date
' @param price      the stock price
'---------------------------------------------------------------------------------------
'
Public Sub InsertStockPrice(StockName As String, StockDate As Date, price As Double)
    Dim BlankRow, source As Range
    Set source = StockPrice.Range("A3:C3")
    Set BlankRow = StockPrice.Range("A" & (LastStockPrice + 1) & ":C" & (LastStockPrice + 1))
    source.Copy
    BlankRow.EntireRow.Insert xlDown, xlFormatFromRightOrBelow
    Application.CutCopyMode = False
    BlankRow.Cells(0, 1).Value = Trim(StockName)
    BlankRow.Cells(0, 2).Value = StockDate
    BlankRow.Cells(0, 3).Value = price
End Sub

'---------------------------------------------------------------------------------------
' Function  : GetUpdateDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return reference to the latest stock price in DataFromWeb or StockPrice
' @param StockName  the stock name
' @param Comment    the output comment (Source : Last update date)
' @return           the reference to latest stock price
' Example           GetLatestStockPrice("AIT") => "=DataFromWeb!F2525"
'---------------------------------------------------------------------------------------
'
'Return Reference to StockPrice or DataFromWeb
Public Function GetLatestStockPrice(StockName As String, ByRef Comment As String) As String
    Dim price As String

    'Compare StockPrice & DataFromWeb Date
    If IsDataFromWebNewer Then
        Comment = "Web : " & DataFromWeb.Range(GetNameRefersTo(Setting.Names("DATE_WEB"))).Value
        price = GetStockCell(StockName)
    End If

    'If price is null get from stockprice
    If price = "" Then
        Comment = "StockPrice : " & StockPrice.Range(GetNameRefersTo(Setting.Names("DATE_STOCKPRICE"))).Value
        price = GetStockPriceCell(StockName)
    End If

    GetLatestStockPrice = price
End Function