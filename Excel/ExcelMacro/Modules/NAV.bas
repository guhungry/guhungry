Attribute VB_Name = "NAV"
'Global Variable
Public LastDataFromWeb As Long
Private DataFromWebSearchColumn As String
Private DataFromWebFirstRow As String

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2012-03-28
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    DataFromWebSearchColumn = GetNameRefersTo(Setting.Names("SEARCH_COLUMN"))
    DataFromWebFirstRow = GetNameRefersTo(Setting.Names("FIRST_ROW"))
    LastDataFromWeb = Utils.GetLastRow(DataFromWeb.Cells, DataFromWebSearchColumn)
End Sub

'---------------------------------------------------------------------------------------
' Function  : FindNAV
' Author    : guhungry
' Date      : 2012-03-28
' Purpose   : Return Range of matched Stock in DataFromWeb, used by GetStockValue and GetStockCell
'---------------------------------------------------------------------------------------
'
Private Function FindNAV(TransactionDate As Date) As Range
    Dim Found As Range
    Dim y As Integer
    Dim sdate As String
    y = Year(TransactionDate) + 543 - 2500
    sdate = Format(TransactionDate, "dd/mm/") & y
    
    Set Found = DataFromWeb.Range(DataFromWebSearchColumn & DataFromWebFirstRow, DataFromWebSearchColumn & LastDataFromWeb)
    Set Found = Found.Find(What:=sdate, After:=Found.Cells(1, 1), LookIn:=xlValues, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Set FindNAV = Found
End Function

'---------------------------------------------------------------------------------------
' Function  : GetValueNAV
' Author    : guhungry
' Date      : 2012-03-28
' @param TransactionDate  the transaction date
' @param PriceType  the transaction type "B"uy, "S"ell or "N"AV
' @return           the NAV
' Example           GetValueNAV("2012-03-28", "B") => 29.5
'---------------------------------------------------------------------------------------
'
Public Function GetValueNAV(TransactionDate As Date, PriceType As String) As Double
    Dim Found As Range
    Set Found = FindNAV(TransactionDate)
    
    If Found Is Nothing Then
        GetValueNAV = -1
    Else
        If (PriceType = "B") Then
            GetValueNAV = Found.Cells(1, 4).Value
        ElseIf (PriceType = "S") Then
            GetValueNAV = Found.Cells(1, 6).Value
        Else
            GetValueNAV = Found.Cells(1, 2).Value
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : GetCellNAV
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return address of matched Stock in DataFromWeb
' @param TransactionDate  the transaction date
' @param PriceType  the transaction type "B"uy, "S"ell or "N"AV
' @return           the address of NAV
' Example           GetCellNAV("2012-03-28", "B") => '=DataFromWeb!F6'
'---------------------------------------------------------------------------------------
'
Public Function GetCellNAV(TransactionDate As Date, TransactionType As String) As String
    Dim Found As Range
    Set Found = FindNAV(TransactionDate)
    
    If Found Is Nothing Then
        GetCellNAV = ""
    Else
        If (PriceType = "B") Then
            GetCellNAV = "=" & Utils.RangeAddress(Found.Cells(1, 4))
        ElseIf (PriceType = "S") Then
            GetCellNAV = "=" & Utils.RangeAddress(Found.Cells(1, 6))
        Else
            GetCellNAV = "=" & Utils.RangeAddress(Found.Cells(1, 2))
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : GetUpdateDate
' Author    : guhungry
' Date      : 2012-03-28
' Purpose   : Return update date of DataFromWeb or StockPrice
' @return           the address of stock price cell
' Example           GetUpdateDate() => '30 June 2010'
'---------------------------------------------------------------------------------------
'
Public Function GetUpdateDate() As String
    GetUpdateDate = TextUtils.ExtractDate(DataFromWeb.Range(GetNameRefersTo(Setting.Names("DATE_WEB"))).Value)
End Function

'---------------------------------------------------------------------------------------
' Function  : GetLastUpdate
' Author    : guhungry
' Date      : 2012-03-28
' Purpose   : Return latest update date from DataFromWeb or StockPrice
' @return           the address of stock price cell
' Example           GetLastUpdate() => 'Last Update 30 Jun 2010 16:59:45'
'---------------------------------------------------------------------------------------
'
Public Function GetLastUpdate() As String
    GetLastUpdate = DataFromWeb.Range(GetNameRefersTo(Setting.Names("DATE_WEB"))).Value
End Function

'---------------------------------------------------------------------------------------
' Function  : GetLatestNAV
' Author    : guhungry
' Date      : 2012-03-28
' Purpose   : Return reference to the latest NAV
' @param Comment    the output comment (Source : Last update date)
' @return           the reference to latest NAV
' Example           GetLatestNAV() => "=DataFromWeb!F2525"
'---------------------------------------------------------------------------------------
'
Public Function GetLatestNAV(ByRef Comment As String) As String
    Dim price As String
    
    Comment = "Web : " & DataFromWeb.Range(GetNameRefersTo(Setting.Names("DATE_WEB"))).Value
    price = DataFromWeb.Range(GetNameRefersTo(Setting.Names("LATEST_NAV"))).Value
    
    GetLatestNAV = price
End Function
