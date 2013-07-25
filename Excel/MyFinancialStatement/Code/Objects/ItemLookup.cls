VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Global Variable
Private LastRow As Long
Private LookupColumn As String

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-14
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize(Lookup As String)
    LastRow = Utils.GetLastRow(ItemLookup.Cells, Lookup)
    LookupColumn = Lookup & "1:" & Lookup & "" & LastRow
End Sub

'---------------------------------------------------------------------------------------
' Function  : FindItem
' Author    : guhungry
' Date      : 2010-07-14
' Purpose   : Return cell of matched Item
' @param Item   the item name
' @return       matched cell
'---------------------------------------------------------------------------------------
'
Private Function FindItem(Item As String) As Range
    Dim Found As Range
    Set Found = ItemLookup.Range(LookupColumn)
    Set Found = Found.Find(What:=Item, After:=Found(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    Set FindItem = Found
End Function

'---------------------------------------------------------------------------------------
' Function  : IsItemExist
' Author    : guhungry
' Date      : 2010-07-14
' Purpose   : Check weather item exist
' @param Item   the item name
' @return       Is Item exists
' Example       IsItemExist("LONG-TERM INVESTMENTS") => True
'---------------------------------------------------------------------------------------
'
Public Function IsItemExist(Item As String) As Boolean
    Dim Found As Range
    
    Set Found = FindItem(Item)
    IsItemExist = Utils.IFF(Found Is Nothing, False, True)
End Function

'---------------------------------------------------------------------------------------
' Function  : Compare
' Author    : guhungry
' Date      : 2010-07-14
' Purpose   : Compare item order
' @param Item1    the 1st item
' @param Item2    the 2nd item
' @return           Is stock exists in StockPrice
' Example           Compare("PREFERED SHARES", "ORDINARY SHARES") => -1
'---------------------------------------------------------------------------------------
'
Public Function Compare(Item1 As String, Item2 As String) As Long
    Dim cell1 As Range, cell2 As Range
    
    Set cell1 = FindItem(Item1)
    Set cell2 = FindItem(Item2)
    Compare = cell1.Row - cell2.Row
End Function
