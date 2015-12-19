Attribute VB_Name = "Utils"
Option Explicit

'---------------------------------------------------------------------------------------
' Function  : GetLastRow
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Get last row number that have data in given column
' @param Cells      the search Range
' @param ColumnName the search column
' @return           the last row number that have data
' Example           GetLastRow(Portfolio.Cells, "A") => 50
'---------------------------------------------------------------------------------------
'
Public Function GetLastRow(Cells As Range, Optional ColumnName As String = "A") As Long
    GetLastRow = Cells.Cells(Cells.Cells.Rows.Count, ColumnName).End(xlUp).Row
End Function

'---------------------------------------------------------------------------------------
' Function  : GetLastColumn
' Author    : guhungry
' Date      : 2010-07-09
' Purpose   : Get last row number that have data in given column
' @param Cells      the search Range
' @param RowNum     the search row
' @return           the last column number that have data
' Example           GetLastColumn(BalanceSheet.Cells, 2) => 50
'---------------------------------------------------------------------------------------
'
Public Function GetLastColumn(Cells As Range, Optional RowNum As Long = 1) As Long
    GetLastColumn = Cells.Cells(RowNum, Cells.Cells.Columns.Count).End(xlToLeft).Column
End Function

'---------------------------------------------------------------------------------------
' Function  : GetNameRefersTo
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Get value of the named variable
' @param value      the named variable
' @return           the value of named variable
' Example           GetNameRefersTo(Setting.Names("DATE_WEB")) => "A2"
'---------------------------------------------------------------------------------------
'
Public Function GetNameRefersTo(Value As Name) As String
    Dim S As String
    Dim HasRef As Boolean
    Dim R As Range
    On Error Resume Next
    Set R = Value.RefersToRange

    HasRef = (err.Number = 0)
    If HasRef = True Then
        S = R.Text
    Else
        S = Value.RefersTo
        If StrComp(Mid(S, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
            ' text constant
            S = Mid(S, 3, Len(S) - 3)
        Else
            ' numeric contant
            S = Mid(S, 2)
        End If
    End If
    GetNameRefersTo = S
End Function

'---------------------------------------------------------------------------------------
' Procedure : DeleteNames
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Delete all Names in Sheet
' @param this       the Names to be deleted
'---------------------------------------------------------------------------------------
'
Public Sub DeleteNames(Value As Names)
    For Each n In Value
        n.Delete
    Next n
End Sub

'---------------------------------------------------------------------------------------
' Function  : RangeAddress
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Get address of given range, use for reference or etc.
' @param Cell      the range
' @return           the address or range
' Example           RangeAddress(Portfolio.Range("B30:G33")) => 'Portfolio!$B$30:$G$33'
'---------------------------------------------------------------------------------------
'
Public Function RangeAddress(Cell As Range) As String
    RangeAddress = "'" & Cell.Parent.Name & "'!" & Replace(Cell.Address, "$", "")
End Function

'---------------------------------------------------------------------------------------
' Function  : ColumnName
' Author    : guhungry
' Date      : 2010-07-14
' Purpose   : Get column name of a cell.
' @param Cell       the cell
' @return           the column name
' Example           ColumnName(Portfolio.Cells(1, 5)) => 'E'
'---------------------------------------------------------------------------------------
'
Public Function ColumnName(Cell As Range) As String
    ColumnName = TextUtils.GetRegExp("[^\d$]+", Cell.Address)
End Function

'---------------------------------------------------------------------------------------
' Function  : IsNumber
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Check if cell is not empty and is number
' @param Cell       the cell to test
' @return           Is not empty and is number
' Example           IsNumber(Portfolio.Range("G30"))
'---------------------------------------------------------------------------------------
'
Public Function IsNumber(Cell As Variant) As Boolean
    IsNumber = (Not (IsEmpty(Cell))) And IsNumeric(Cell)
End Function

'---------------------------------------------------------------------------------------
' Function  : IFF
' Author    : guhungry
' Date      : 2010-07-08
' Purpose   : IF in macro style
' @param Expression     the logic expression
' @param ValueTrue      the value if true
' @param ValueFalse     the value if false
' @return           returns ValueTrue if true else returns ValueFalse
' Example           IFF(1=1, "True", "False") => "True"
'---------------------------------------------------------------------------------------
'
Public Function IFF(Expression As Boolean, ValueTrue As Variant, ValueFalse As Variant) As Variant
    If Expression Then
        IFF = ValueTrue
    Else
        IFF = ValueFalse
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : FitComment
' Author    : guhungry
' Date      : 2010-07-30
' Purpose   : Dear god please fit comment for me
'---------------------------------------------------------------------------------------
'
Public Sub FitComment()
    ActiveCell.Comment.Shape.TextFrame.AutoSize = True
End Sub