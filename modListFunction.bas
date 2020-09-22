Attribute VB_Name = "modListFunction"
Option Explicit

'Thanks to Macromedia ColdFusion who tought me the concept of LIST thing

Function ListGetAt(List As String, Element As Long, Optional Delimeter As String = ",", Optional ErrorDialog As Boolean = True) As String
'if the List contains less elements than specified in Element parameter, there are two responses;
'   ErrorDialog = True      Raise an error (Default)
'   ErrorDialog = False     Don't raise any error and return NULL

ListGetAt = vbNull
If Element > ListLen(List, Delimeter) Then
    If ErrorDialog Then MsgBox "The List parameter contains less element(s) than specified in Element parameter!", vbCritical, "Error!"
    Exit Function
End If

Dim a As Long, Temp As String
If Element = 1 Then
    If ListLen(List, Delimeter) = 1 Then ListGetAt = List Else ListGetAt = Left(List, InStr(List, Delimeter) - 1)
Else
    Temp = List
    For a = 1 To Element - 1
        Temp = Mid(Temp, InStr(Temp, Delimeter) + 1)
    Next
    If InStr(Temp, Delimeter) = 0 Then ListGetAt = Temp Else ListGetAt = Left(Temp, InStr(Temp, Delimeter) - 1)
End If
End Function

Function ListLen(List As String, Optional Delimeter As String = ",") As Long
If List = "" Then Exit Function

Dim Temp As String, FoundCount As Long
Temp = List
While InStr(Temp, Delimeter) <> 0
    DoEvents
    If InStr(Temp, Delimeter) > 0 Then
        FoundCount = FoundCount + 1
        Temp = Mid(Temp, InStr(Temp, Delimeter) + 1)
    End If
Wend
ListLen = FoundCount + 1
End Function

Function ListFind(List As String, SearchText As String, Optional Delimeter As String = ",", Optional CaseSensitive As Boolean = False) As Long
ListFind = 0
If ListLen(List, Delimeter) = 0 Then Exit Function

Dim a As Long, Found As Boolean

For a = 1 To ListLen(List, Delimeter)
    If CaseSensitive Then
        If SearchText = ListGetAt(List, a, Delimeter) Then Found = True
    Else
        If LCase(SearchText) = ListGetAt(LCase(List), a, Delimeter) Then Found = True
    End If
    If Found Then Exit For
Next
If Found Then ListFind = a
End Function

Function ListDeleteAt(List As String, Element As Long, Optional Delimeter As String = ",", Optional ErrorDialog As Boolean = True) As String
Dim a As Long, b As Long, ResultList As String

If Element > ListLen(List, Delimeter) Then
    If ErrorDialog Then MsgBox "The List parameter contains less element(s) than specified in Element parameter!", vbCritical, "Error!"
    Exit Function
End If

For a = 1 To Element - 1
    ResultList = ListAppend(ResultList, ListGetAt(List, a, Delimeter), ",")
Next
If ListLen(List, ",") > Element Then
    For a = Element + 1 To ListLen(List, Delimeter)
        ResultList = ListAppend(ResultList, ListGetAt(List, a, Delimeter), ",")
    Next
End If

ListDeleteAt = ResultList
End Function

Function ListAppend(List As String, Value As String, Optional Delimeter As String = ",") As String
If ListLen(List, Delimeter) = 0 Then ListAppend = Value Else ListAppend = List & Delimeter & Value
End Function

Function ListInsertAt(List As String, Value As String, Position As Long, Delimeter As String, Optional ErrorDialog As Boolean = True) As String
If ListLen(List, Delimeter) < Position - 1 Then
    If ErrorDialog Then MsgBox "The List parameter contains less element(s) than specified in Position parameter!", vbCritical, "Error!"
    Exit Function
End If

Dim a As Long, ResultList As String
For a = 1 To Position - 1
    ResultList = ListAppend(ResultList, ListGetAt(List, a, Delimeter), Delimeter)
Next
ResultList = ListAppend(ResultList, Value, Delimeter)
For a = Position To ListLen(List, Delimeter)
    ResultList = ListAppend(ResultList, ListGetAt(List, a, Delimeter), Delimeter)
Next
ListInsertAt = ResultList
End Function
