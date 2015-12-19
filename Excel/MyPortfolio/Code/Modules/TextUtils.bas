Attribute VB_Name = "TextUtils"
Option Explicit

'---------------------------------------------------------------------------------------
' Function  : GetRegExp
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Search subject for a match to the regular expression given in pattern.
' @param pattern  the regular expression pattern
' @param subject  the input string
' @return           the matched substring
' Example           GetRegExp("\d{1,2} [A-Za-z]+ \d{4}", "Last Update 30 Jun 2010 16:59:45") => "30 Jun 2010"
'---------------------------------------------------------------------------------------
'
Public Function GetRegExp(myPattern As String, myString As String, Optional IsIgnoreCase As Boolean = True, Optional IsGlobal As Boolean = True) As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches As MatchCollection
    Dim RetStr As String

    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.pattern = myPattern

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = IsIgnoreCase

    'Set global applicability.
    objRegExp.Global = IsGlobal

    'Test whether the String can be compared.
    If (objRegExp.test(myString) = True) Then

        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   ' Execute search.

        For Each objMatch In colMatches   ' Iterate Matches collection.
            RetStr = RetStr & objMatch.Value
        Next
    Else
        RetStr = ""
    End If
    GetRegExp = RetStr
End Function

'---------------------------------------------------------------------------------------
' Function  : TestRegExp
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Test subject for a match to the regular expression given in pattern.
' @param pattern  the regular expression pattern
' @param subject  the input string
' @return           Is subject match to pattern ('True', 'False')
' Example           TestRegExp("\d{1,2} [A-Za-z]+ \d{4}", "Last Update 30 Jun 2010 16:59:45") => True
'---------------------------------------------------------------------------------------
'
Public Function TestRegExp(pattern As String, subject As String) As Boolean
    'Create objects.
    Dim objRegExp As RegExp

    ' Create a regular expression object.
    Set objRegExp = New RegExp

    'Set the pattern by using the Pattern property.
    objRegExp.pattern = pattern

    ' Set Case Insensitivity.
    objRegExp.IgnoreCase = True

    'Set global applicability.
    objRegExp.Global = True

    'Test whether the String can be compared.
    TestRegExp = objRegExp.test(subject)
End Function

'---------------------------------------------------------------------------------------
' Function  : ExtractDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Return long date matched in input string
' @param subject  the input string
' @return           the matched date
' Example           ExtractDate("Last Update 30 Jun 2010 16:59:45") => '30 Jun 2010'
'---------------------------------------------------------------------------------------
'2015-12-18 Remove using Regular Expression which not support by Mac
Public Function ExtractDate(subject As String) As Date
    subject = Replace(subject, "* Last Update ", "")
    subject = Replace(subject, "* ข้อมูลล่าสุด ", "")
    subject = Left(subject, Len(subject) - 9)

    ExtractDate = DateValue(subject)
End Function