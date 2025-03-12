Option Explicit

Private Function GetRegex(Optional pattern As String = "", Optional ignoreCase As Boolean = True, Optional globalFlag As Boolean = True) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.IgnoreCase = ignoreCase
    regex.Global = globalFlag
    Set GetRegex = regex
End Function

Public Function ReSearch(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Object
    Dim regex As Object
    Set regex = GetRegex(pattern, ignoreCase, False)
    If regex.Test(text) Then
        Set ReSearch = regex.Execute(text)(0)
    Else
        Set ReSearch = Nothing
    End If
End Function

Public Function ReMatch(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Object
    Dim regex As Object
    Set regex = GetRegex("^" & pattern, ignoreCase, False)
    If regex.Test(text) Then
        Set ReMatch = regex.Execute(text)(0)
    Else
        Set ReMatch = Nothing
    End If
End Function

Public Function ReFindAll(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Collection
    Dim regex As Object, matches As Object, matchItem As Variant
    Dim results As New Collection
    Set regex = GetRegex(pattern, ignoreCase, True)
    Set matches = regex.Execute(text)
    For Each matchItem In matches
        results.Add matchItem.Value
    Next matchItem
    Set ReFindAll = results
End Function

Public Function ReSub(pattern As String, replacement As String, text As String, Optional ignoreCase As Boolean = True) As String
    Dim regex As Object
    Set regex = GetRegex(pattern, ignoreCase, True)
    ReSub = regex.Replace(text, replacement)
End Function

Public Function ReSplit(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Collection
    Dim regex As Object, matches As Object
    Dim result As New Collection
    Dim lastEnd As Long, currentPos As Long
    Dim i As Long
    Set regex = GetRegex(pattern, ignoreCase, True)
    Set matches = regex.Execute(text)
    lastEnd = 1
    For i = 0 To matches.Count - 1
        currentPos = matches(i).FirstIndex + 1
        result.Add Mid(text, lastEnd, currentPos - lastEnd)
        lastEnd = currentPos + matches(i).Length
    Next i
    result.Add Mid(text, lastEnd)
    Set ReSplit = result
End Function

Public Function ReGroup(matchObj As Object, Optional groupNumber As Integer = 0) As String
    Dim subMatches As Variant
    If matchObj Is Nothing Then
        ReGroup = ""
        Exit Function
    End If
    If groupNumber = 0 Then
        ReGroup = matchObj.Value
    Else
        subMatches = matchObj.SubMatches
        If groupNumber - 1 <= UBound(subMatches) Then
            ReGroup = subMatches(groupNumber - 1)
        Else
            ReGroup = ""
        End If
    End If
End Function

Public Function ReGroups(matchObj As Object) As Collection
    Dim groups As New Collection
    Dim subMatches As Variant
    Dim i As Integer
    If matchObj Is Nothing Then
        Set ReGroups = groups
        Exit Function
    End If
    subMatches = matchObj.SubMatches
    For i = LBound(subMatches) To UBound(subMatches)
        groups.Add subMatches(i)
    Next i
    Set ReGroups = groups
End Function

Public Function ReSearchObj(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Object
    Set ReSearchObj = ReSearch(pattern, text, ignoreCase)
End Function

Public Function ReSearchBool(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Boolean
    Dim regex As Object
    Set regex = GetRegex(pattern, ignoreCase, False)
    ReSearchBool = regex.Test(text)
End Function

Public Function ReMatchObj(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Object
    Set ReMatchObj = ReMatch(pattern, text, ignoreCase)
End Function

Public Function ReMatchStr(pattern As String, text As String, Optional ignoreCase As Boolean = True) As String
    Dim matchObj As Object
    Set matchObj = ReMatch(pattern, text, ignoreCase)
    If Not matchObj Is Nothing Then
        ReMatchStr = matchObj.Value
    Else
        ReMatchStr = ""
    End If
End Function