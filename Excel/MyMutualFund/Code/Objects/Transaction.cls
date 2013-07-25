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
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    info.Init Transaction.Cells, 10, "B"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionAdd
' Author    : guhungry
' Date      : 2012-03-28
' Purpose   : Add stock transaction (Buy, Sell, Dividend) to Transaction Sheet, Check error, add record and styling
'---------------------------------------------------------------------------------------
'
Sub btnTransactionAdd()
    'Define variables
    Dim err As String
    Dim source As Range, target As Range
    err = ""
    Set source = Transaction.Range("B6")
    Set target = Transaction.Range("B11")
    
    NAV.Initialize
    
    'Validate Data
    If IsEmpty(source.Cells(1, 1).Value) Then
        source.Cells(1, 1).Value = "=Today()"
    End If
    If Not (IsDate(source.Cells(1, 1).Value)) Then
        err = err & "Date is not valid." & vbNewLine
    End If
    'Validate Transaction Type (Buy, Sale, Dividend)
    If Not (TextUtils.TestRegExp("^[BSD]$", source.Cells(1, 2).Value)) Then
        source.Cells(1, 2).Value = "B"
    End If
    source.Cells(1, 2).Value = UCase(source.Cells(1, 2).Value)
    'Validate Value, Quantity, Price (NAV)
    If (source.Cells(1, 2).Value = "D") Then
        'Empty the Value
        source.Cells(1, 3).Value = ""
        'Validate NAV
        If Not (Utils.IsNumber(source.Cells(1, 5).Value)) Then
            err = err & "Please insert price (dividend)." & vbNewLine
        End If
        'Validate Quantity
        If Not (Utils.IsNumber(source.Cells(1, 4).Value)) Then
            err = err & "Please insert quantity." & vbNewLine
        End If
    Else
        ' Validate Value
        If Not (Utils.IsNumber(source.Cells(1, 3).Value)) Then
            err = err & "Please insert value." & vbNewLine
        Else
            source.Cells(1, 3).Value = Abs(source.Cells(1, 3).Value)
        End If
        'Fill With Latest NAV
        source.Cells(1, 5).Value = NAV.GetValueNAV(source.Cells(1, 1).Value, source.Cells(1, 2).Value)
        If (source.Cells(1, 5).Value = -1) Then
            source.Cells(1, 5).Value = ""
            err = err & "Please update data from web." & vbNewLine
        End If
        If err = "" Then
            'Calculate Quantity
            If source.Cells(1, 2).Value = "S" Then
                source.Cells(1, 4).Value = Application.RoundUp(source.Cells(1, 3).Value / source.Cells(1, 5).Value, 5)
            Else
                source.Cells(1, 4).Value = Application.RoundUp(source.Cells(1, 3).Value / source.Cells(1, 5).Value, 6)
            End If
            source.Cells(1, 4).Value = Application.RoundDown(source.Cells(1, 4).Value, 4)
            'Set NAV Price
            source.Cells(1, 5).Value = NAV.GetValueNAV(source.Cells(1, 1).Value, "N")
        End If
    End If
    If err <> "" Then
        MsgBox err
        Exit Sub
    End If
    
    'Add Data
    target.EntireRow.Insert xlDown, xlFormatFromRightOrBelow
    target.Cells(0, 1).Value = source.Cells(1, 1).Value
    If source.Cells(1, 2).Value = "B" Then
        target.Cells(0, 3).Value = source.Cells(1, 4).Value
        target.Cells(0, 2).Value = -source.Cells(1, 3).Value
    ElseIf source.Cells(1, 2).Value = "S" Then
        target.Cells(0, 4).Value = source.Cells(1, 4).Value
        target.Cells(0, 2).Value = source.Cells(1, 3).Value
    Else
        target.Cells(0, 5).Value = source.Cells(1, 4).Value
    End If
    target.Cells(0, 6).Value = source.Cells(1, 5).Value
    target.Cells(0, 7).FormulaR1C1 = "=ROUND(IF(RC[-4]>0, -1*RC[-4]*RC[-1], IF(RC[-3]>0, RC[-3]*RC[-1], RC[-2]*RC[-1])), 2)"
    If (source.Cells(1, 2).Value <> "D") Then
        target.Cells(0, 8).FormulaR1C1 = "=(RC[-6]-RC[-1])"
    Else
        target.Cells(0, 2).Value = target.Cells(0, 7).Value
    End If
    target.Cells(0, 9).Value = source.Cells(1, 6).Value
    
    'Styling
    Initialize
    TransactionStyleAdd
    info.RestoreActiveCell
    btnTransactionSortByDate
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionSortByDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Sort stock transaction by date
'---------------------------------------------------------------------------------------
'
Sub btnTransactionSortByDate()
    Dim myActive As String
    TransactionSort "B", xlDescending
    
    'Make Report
    Application.ScreenUpdating = False
    Report.Initialize
    myActive = Report.ActCell
    Report.ReportClearData
    Report.ReportMake
    Report.Range(myActive).Select
    Transaction.Activate
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TransactionStyleAdd
' Author    : guhungry
' Date      : 2010-07-06
' Purpose   : Copy styling template to report data
'---------------------------------------------------------------------------------------
'
Private Sub TransactionStyleAdd()
    Dim source As Range, target As Range
    
    If info.LastRow > 10 Then
        Application.ScreenUpdating = False
        Set source = Transaction.Range("B9:J9")
        Set target = Transaction.Range("B11:J" & info.LastRow)
        
        source.Copy
        target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionSortByTicker
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Based stock transaction sort procedure
' @param Column1    the 1st column to sort
' @param Order1     the 1st sort order
'---------------------------------------------------------------------------------------
'
Private Sub TransactionSort(Column1 As String, Order1 As XlSortOrder)
    Dim target As Range

    Initialize
    Application.ScreenUpdating = False
    If info.LastRow > 10 Then
        Set target = Transaction.Range("B11:J" & info.LastRow)
        
        'Sort
        Transaction.Sort.SortFields.Clear
        Transaction.Sort.SortFields.Add Key:=Range(Column1 & "11:" & Column1 & info.LastRow), SortOn:=xlSortOnValues, Order:=Order1, DataOption:=xlSortNormal
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
