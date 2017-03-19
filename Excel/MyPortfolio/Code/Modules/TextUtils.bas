Attribute VB_Name = "TextUtils"
Option Explicit

'---------------------------------------------------------------------------------------
' Function : GetRegExp
' Author : guhungry
' Date : 2010-07-01
' Purpose : Search subject for a match to the regular expression given in pattern.
' @param pattern the regular expression pattern
' @param subject the input string
' @return the matched substring
' Example GetRegExp("\d{1,2} [A-Za-z]+ \d{4}", "Last Update 30 Jun 2010 16:59:45") => "30 Jun 2010"
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
        Set colMatches = objRegExp.Execute(myString) ' Execute search.

        For Each objMatch In colMatches ' Iterate Matches collection.
            RetStr = RetStr & objMatch.Value
        Next
    Else
        RetStr = ""
    End If
    GetRegExp = RetStr
End Function

'---------------------------------------------------------------------------------------
' Function : TestRegExp
' Author : guhungry
' Date : 2010-07-01
' Purpose : Test subject for a match to the regular expression given in pattern.
' @param pattern the regular expression pattern
' @param subject the input string
' @return Is subject match to pattern ('True', 'False')
' Example TestRegExp("\d{1,2} [A-Za-z]+ \d{4}", "Last Update 30 Jun 2010 16:59:45") => True
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
' Function : ExtractDate
' Author : guhungry
' Date : 2010-07-01
' Purpose : Return long date matched in input string
' @param subject the input string
' @return the matched date
' Example ExtractDate("Last Update 30 Jun 2010 16:59:45") => '30 Jun 2010'
'---------------------------------------------------------------------------------------
'
Public Function ExtractDate(subject As String) As Date
    Dim val As String
    val = TextUtils.GetRegExp("\d{1,2} [A-Za-z]+ \d{4}", subject)
    If val = "" Then
        val = TextUtils.GetRegExp("\d{1,2}/\d{1,2}/\d{4}", subject)
        val = FormatDate(val)
    End If
    ExtractDate = DateValue(val)
End Function

'---------------------------------------------------------------------------------------
' Function : FormatDate
' Author : guhungry
' Date : 2017-03-13
' Purpose : Return date string in dd MMM yyyy format
' @param subject the date string in dd/MM/yyyy format
' @return the formatted date
' Example FormatDate("13/03/2017") => '13 Mar 2017'
'---------------------------------------------------------------------------------------
'
Public Function FormatDate(subject As String) As Date
    subject = Replace(subject, "/01/", " Jan ")
    subject = Replace(subject, "/02/", " Feb ")
    subject = Replace(subject, "/03/", " Mar ")
    subject = Replace(subject, "/04/", " Apr ")
    subject = Replace(subject, "/05/", " May ")
    subject = Replace(subject, "/06/", " Jun ")
    subject = Replace(subject, "/07/", " Jul ")
    subject = Replace(subject, "/08/", " Aug ")
    subject = Replace(subject, "/09/", " Sep ")
    subject = Replace(subject, "/10/", " Oct ")
    subject = Replace(subject, "/11/", " Nov ")
    FormatDate = Replace(subject, "/12/", " Dec ")
End Function
