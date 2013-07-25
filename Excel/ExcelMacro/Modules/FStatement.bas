Attribute VB_Name = "FStatement"
Option Explicit
'Financial Statement Module
Private TargetSheet As Range
Private Row As Long

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-13
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize(LookupColumn As String)
    Row = 2
    Set TargetSheet = Temp.Cells
    ItemLookup.Initialize LookupColumn
End Sub

'---------------------------------------------------------------------------------------
' Procedure : MergeStatement
' Author    : guhungry
' Date      : 2010-07-08
' Purpose   : Merge Financial Statement from Source2 into Source1
' @param Source1    the statement 1
' @param Source2    the statement 2
' @restriction      Newer statement's last row must overlapped with older's first row
'---------------------------------------------------------------------------------------
'
Public Sub MergeStatement(Source1 As Range, Source2 As Range, LookupColumn As String)
    Dim NewSource As Range, OldSource As Range
    Dim myTarget As Range
    
    Initialize LookupColumn
    If Source1(1, 2).Value = Source2(1, 2).Value Then
        MsgBox "Can not merge statement from same year", vbCritical
        Exit Sub
    ElseIf Source1(1, 2).Value > Source2(1, 2).Value Then
        Set NewSource = Source1
        Set OldSource = Source2
    Else
        Set NewSource = Source2
        Set OldSource = Source1
    End If
    
    'Generate Data to Temp (Using Merge Sort Algorithm)
    Dim oldSheet As New SheetInfo, newSheet As New SheetInfo, targetInfo As New SheetInfo
    oldSheet.Init OldSource.Cells
    newSheet.Init NewSource.Cells
    
    'Error if data is not overlapped
    If oldSheet.LastRow > 1 And OldSource(1, 2).Value <> NewSource(1, newSheet.LastCol).Value Then
        MsgBox "Can not merge statement. Data is not overlapping", vbCritical
        Exit Sub
    End If
    
    Do
        Set OldSource = oldSheet.CurCell
        Set NewSource = newSheet.CurCell
        
        'No data left
        If NewSource Is Nothing And OldSource Is Nothing Then
            Exit Do
        'Only new data left
        ElseIf OldSource Is Nothing Then
            PrintAllNew oldSheet, newSheet
            Exit Do
        'Only old data left
        ElseIf NewSource Is Nothing Then
            PrintAllOld oldSheet, newSheet
            Exit Do
        'Record matched
        ElseIf OldSource.Value = NewSource.Value And OldSource(1, 2).Value = NewSource(1, newSheet.LastCol).Value Then
            PrintAll oldSheet, newSheet
        'Print only old record
        ElseIf OldSource(1, 2).Value = 0 And NewSource(1, newSheet.LastCol).Value <> 0 Then
            PrintOld oldSheet, newSheet
        'Print only new record
        ElseIf OldSource(1, 2).Value <> 0 And NewSource(1, newSheet.LastCol).Value = 0 Then
            PrintNew oldSheet, newSheet
        ElseIf ItemLookup.IsItemExist(NewSource.Value) And ItemLookup.IsItemExist(OldSource.Value) Then
            If ItemLookup.Compare(OldSource.Value, NewSource.Value) > 0 Then
                PrintNew oldSheet, newSheet
            Else
                PrintOld oldSheet, newSheet
            End If
        Else
            MsgBox "FixMe : FStatement.MergeStatement " & Row & vbNewLine & _
                "[" & oldSheet.CurCell.Value & ":" & oldSheet.CurCell(1, 2).Value & "]" & vbNewLine & _
                "[" & newSheet.CurCell.Value & ":" & newSheet.CurCell(1, newSheet.LastCol).Value & "]"
            Exit Sub
        End If
    Loop
    
    'Clear Data Between Old & New
    TargetSheet.Range(Utils.ColumnName(TargetSheet(1, newSheet.LastCol + 1)) & ":" & Utils.ColumnName(TargetSheet(1, newSheet.LastCol + 2))).Delete
    
    'Styling
    targetInfo.Init TargetSheet.Cells, 2
    Set myTarget = targetInfo.GetRange("A2")
    TargetSheet(1, 1).Copy
    myTarget.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    'Copy Data Back to Target
    myTarget.Cells.Copy
    Source1.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    myTarget.Delete xlShiftUp
    Source2.Cells.Delete xlShiftUp
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintNew
' Author    : guhungry
' Date      : 2010-07-12
' Purpose   : Print only new source, old source is 0
' @param OldSource    the old source
' @param NewSource    the new source
'---------------------------------------------------------------------------------------
'
Private Sub PrintNew(OldSource As SheetInfo, ByRef NewSource As SheetInfo)
    PrintData NewSource, 2
    PrintZero OldSource, NewSource.LastCol + 2
    NewSource.CurRow = NewSource.CurRow + 1
    Row = Row + 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintOld
' Author    : guhungry
' Date      : 2010-07-12
' Purpose   : Print only old source, new source is 0
' @param OldSource    the old source
' @param NewSource    the new source
'---------------------------------------------------------------------------------------
'
Private Sub PrintOld(ByRef OldSource As SheetInfo, NewSource As SheetInfo)
    PrintZero NewSource, 2
    PrintData OldSource, NewSource.LastCol + 2
    OldSource.CurRow = OldSource.CurRow + 1
    Row = Row + 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintAll
' Author    : guhungry
' Date      : 2010-07-12
' Purpose   : Print both new source and old source
' @param OldSource    the old source
' @param NewSource    the new source
'---------------------------------------------------------------------------------------
'
Private Sub PrintAll(ByRef OldSource As SheetInfo, ByRef NewSource As SheetInfo)
    PrintData NewSource, 2
    PrintData OldSource, NewSource.LastCol + 2
    OldSource.CurRow = OldSource.CurRow + 1
    NewSource.CurRow = NewSource.CurRow + 1
    Row = Row + 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintAllNew
' Author    : guhungry
' Date      : 2010-07-12
' Purpose   : Print only new source from current row to last row
' @param OldSource    the old source
' @param NewSource    the new source
'---------------------------------------------------------------------------------------
'
Private Sub PrintAllNew(OldSource As SheetInfo, ByRef NewSource As SheetInfo)
    Dim i As Long
    For i = NewSource.CurRow To NewSource.LastRow
        PrintNew OldSource, NewSource
        NewSource.CurRow = i + 1
    Next i
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintAllOld
' Author    : guhungry
' Date      : 2010-07-12
' Purpose   : Print only old source from current row to last row
' @param OldSource    the old source
' @param NewSource    the new source
'---------------------------------------------------------------------------------------
'
Private Sub PrintAllOld(ByRef OldSource As SheetInfo, NewSource As SheetInfo)
    Dim i As Long
    For i = OldSource.CurRow To OldSource.LastRow
        PrintOld OldSource, NewSource
        OldSource.CurRow = i + 1
    Next i
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintData
' Author    : guhungry
' Date      : 2010-07-13
' Purpose   : Print Data at target cell
' @param Source    the data source
' @param CellAt    the target column
'---------------------------------------------------------------------------------------
'
Private Sub PrintData(Source As SheetInfo, CellAt As Long)
    Dim TargetCell As Range, SourceCell As Range
    Dim i As Long
    Set TargetCell = TargetSheet(Row, 1)
    Set SourceCell = Source.CurCell
    
    TargetCell.Value = SourceCell.Value
    For i = 2 To Source.LastCol
        TargetCell(1, CellAt + i - 2).Value = SourceCell(1, i).Value
    Next i
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintZero
' Author    : guhungry
' Date      : 2010-07-13
' Purpose   : Print Zero at target cell
' @param Source    the data source
' @param CellAt    the target column
'---------------------------------------------------------------------------------------
'
Private Sub PrintZero(Source As SheetInfo, CellAt As Long)
    Dim TargetCell As Range
    Dim i As Long
    Set TargetCell = TargetSheet(Row, 1)
    
    For i = 2 To Source.LastCol
        TargetCell(1, CellAt + i - 2).Value = 0
    Next i
End Sub
