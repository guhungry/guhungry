VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report"
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
    info.Init Report.Cells, 3
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportMake
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Generate stock transaction report
'---------------------------------------------------------------------------------------
'
Public Sub ReportMake()
    Dim curDate As Date
    Dim myComment As String
    Dim myCell As Range, endCell As Range, tran As Range, target As Range
    Dim LastRow As Long, EndRow As Long, StartRow As Long, NumStock As Long
    
    'Initialize Variables
    Report.EnableCalculation = False
    NAV.Initialize
    Transaction.Initialize
    Report.Activate
    Set myCell = Report.Cells(6, 1)
    LastRow = 11
    StartRow = LastRow
    Do While LastRow <= Transaction.LastRow
        Set tran = Transaction.Range("C" & LastRow)
        curDate = tran.Cells(1, 0).Value
        
        'Set Header in New Row
        If StartRow = LastRow Then
            NumStock = 0
            myCell.Cells(1, 3).Value = Trim(TextUtils.GetRegExp("^[^(]+", DataFromWeb.Range(GetNameRefersTo(Setting.Names("FUND_NAME"))).Value))
            Set myCell = myCell.Cells(2, 1)
        End If
        
        'Find EndRow
        NumStock = NumStock + 1
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

        'Insert Data
        myCell.Cells(1, 1).Value = curDate
        myCell.Cells(1, 4).Value = "=SUM(Transaction!C" & EndRow & ":C" & LastRow & ")"
        'Buy
        myCell.Cells(1, 11).Value = "=SUMIF(Transaction!D" & EndRow & ":D" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
        myCell.Cells(1, 13).Value = "=-SUMIF(Transaction!D" & EndRow & ":D" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & LastRow & ")"
        myCell.Cells(1, 12).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1]," & Chr$(34) & Chr$(34) & ")"
        'Sell
        myCell.Cells(1, 14).Value = "=SUMIF(Transaction!E" & EndRow & ":E" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
        myCell.Cells(1, 16).Value = "=SUMIF(Transaction!E" & EndRow & ":E" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & LastRow & ")"
        myCell.Cells(1, 15).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1]," & Chr$(34) & Chr$(34) & ")"
        'Dividend
        myCell.Cells(1, 18).Value = "=SUMIF(Transaction!F" & EndRow & ":F" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & LastRow & ")"
        'Remain Stock
        myCell.Cells(1, 5).FormulaR1C1 = "=RC[6]-RC[9]+R[1]C"
        'Commission
        myCell.Cells(1, 17).Value = "=-SUM(Transaction!I" & EndRow & ":I" & LastRow & ")"
        
        'Insert Header in Last Row
        If EndRow = Transaction.LastRow Then
            Set myCell = myCell.Cells(2, 1)
            Set target = myCell.Cells(-NumStock, 1)
            target.Cells(1, 4).Value = "=SUM(Transaction!C" & EndRow & ":C" & StartRow & ")"
            'Buy
            target.Cells(1, 11).Value = "=SUMIF(Transaction!D" & EndRow & ":D" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
            target.Cells(1, 13).Value = "=-SUMIF(Transaction!D" & EndRow & ":D" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & StartRow & ")"
            target.Cells(1, 12).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1]," & Chr$(34) & Chr$(34) & ")"
            'Sell
            target.Cells(1, 14).Value = "=SUMIF(Transaction!E" & EndRow & ":E" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
            target.Cells(1, 16).Value = "=SUMIF(Transaction!E" & EndRow & ":E" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & StartRow & ")"
            target.Cells(1, 15).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1], " & Chr$(34) & Chr$(34) & ")"
            'Dividend
            target.Cells(1, 18).Value = "=SUMIF(Transaction!F" & EndRow & ":F" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & StartRow & ")"
            'Remain Stock
            target.Cells(1, 5).FormulaR1C1 = "=RC[6]-RC[9]"
            'Commission
            target.Cells(1, 17).Value = "=-SUM(Transaction!I" & EndRow & ":I" & StartRow & ")"

            Report.EnableCalculation = True
            If target.Cells(1, 5).Value > 0 Then
                target.Cells(1, 6).Value = "=-RC[-2]/RC[-1]"
                'Get Latest Stock Price
                target.Cells(1, 7).Value = NAV.GetLatestNAV(myComment)
                target.Cells(1, 7).AddComment myComment
                target.Cells(1, 7).Comment.Shape.TextFrame.AutoSize = True
                target.Cells(1, 8).Value = "=RC[-1]*RC[-3]"
                target.Cells(1, 9).Value = "=RC[-1]+RC[-5]"
                target.Cells(1, 10).Value = "=IF(RC[-6]>0,RC[-1]*0.01,-RC[-1]/RC[-6])"
            End If
            Report.EnableCalculation = False
        End If
        
        Set myCell = myCell.Cells(2, 1)
        LastRow = EndRow + 1
    Loop

    'Main Header
    Set myCell = Report.Range("D5")
    LastRow = Transaction.LastRow
    If LastRow > 10 Then
        myCell.Cells(1, 1).Value = "=SUM(Transaction!C11:C" & LastRow & ")"
        myCell.Cells(1, 8).Value = "=-SUMIF(Transaction!D11:D" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C11:C" & LastRow & ")"
        myCell.Cells(1, 11).Value = "=SUMIF(Transaction!E11:E" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C11:C" & LastRow & ")"
        myCell.Cells(1, 14).Value = "=SUM(Transaction!I11:I" & LastRow & ")"
        myCell.Cells(1, 15).Value = "=SUMIF(Transaction!F11:F" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C11:C" & LastRow & ")"
        LastRow = Utils.GetLastRow(Report.Cells)
        myCell.Cells(1, 5).Value = "=SUM(H6:H" & LastRow & ")"
        myCell.Cells(1, 6).Value = "=SUM(I6:I" & LastRow & ")"
        myCell.Cells(1, 7).Value = "=IF(RC[-6]>0,RC[-1]*0.01,-RC[-1]/RC[-6])"
    End If
    
    'Styling
    Dim myActive As String
    myActive = info.ActCell
    Initialize
    ReportStyleAdd
    Report.Range(myActive).Select
    Report.EnableCalculation = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportClearData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Clear report data
'---------------------------------------------------------------------------------------
'
Public Sub ReportClearData()
    'Clear Header
    Dim Header As Range
    Set Header = Report.Range("D5")
    Header.Value = 0
    Header.Cells(1, 5).Value = 0
    Header.Cells(1, 6).Value = 0
    Header.Cells(1, 7).Value = 0
    Header.Cells(1, 8).Value = 0
    Header.Cells(1, 11).Value = 0
    Header.Cells(1, 14).Value = 0
    Header.Cells(1, 15).Value = 0

    'Clear Data
    If LastRow > 5 Then
        Report.Range("A6:R" & LastRow).Delete Shift:=xlUp
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportStyleAdd
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Copy styling template to report data
'---------------------------------------------------------------------------------------
'
Private Sub ReportStyleAdd()
    Dim source As Range, target As Range
    
    If LastRow > 5 Then
        Set source = Report.Range("A2:R2")
        Set target = Report.Range("A6:R" & LastRow)
        
        source.Copy
        target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
End Sub

