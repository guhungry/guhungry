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
'             Fix Fucked Up Web Data - Cause of Listed or Delisted stock(s).
'---------------------------------------------------------------------------------------
'
Sub btnResetWebData()
    'Define Variables
    Dim LastRow As Integer
    Dim DataFromWebSearchColumn As String
    DataFromWebSearchColumn = GetNameRefersTo(Setting.Names("SEARCH_COLUMN"))
    LastRow = Utils.GetLastRow(DataFromWeb.Cells, DataFromWebSearchColumn) + 2

    Application.ScreenUpdating = False
    'Clear Data
    Range("A1:L" & LastRow).Delete Shift:=xlUp
    LastRow = 1
    'SET Main
    SetDataFromWeb "SetMain", GetNameRefersTo(Setting.Names("URL_SETMAIN")), GetNameRefersTo(Setting.Names("TAB_SETMAIN")), "A" & LastRow

    Transaction.btnTransactionSortByDate
    DataFromWeb.Activate
    Application.ScreenUpdating = True
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
        .RefreshOnFileOpen = True
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
        .WebDisableDateRecognition = True
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

