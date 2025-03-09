Option Explicit

' 内部で RegExp オブジェクトを生成するヘルパー関数
Private Function getRegex(Optional pattern As String = "", Optional ignoreCase As Boolean = True, Optional globalFlag As Boolean = True) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.IgnoreCase = ignoreCase
    regex.Global = globalFlag
    Set getRegex = regex
End Function

' re.search: 対象文字列内から最初のマッチを返す（マッチがなければ Nothing）
Public Function re_search(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Object
    Dim regex As Object
    Set regex = getRegex(pattern, ignoreCase, False) ' 最初の1件のみ取得
    If regex.Test(text) Then
        Set re_search = regex.Execute(text)(0)
    Else
        Set re_search = Nothing
    End If
End Function

' re.match: 文字列の先頭からパターンにマッチするかチェック（マッチがなければ Nothing）
Public Function re_match(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Object
    Dim regex As Object
    ' 文字列の先頭からのマッチを確認するため、パターンの先頭に "^" を追加
    Set regex = getRegex("^" & pattern, ignoreCase, False)
    If regex.Test(text) Then
        Set re_match = regex.Execute(text)(0)
    Else
        Set re_match = Nothing
    End If
End Function

' re.findall: 対象文字列内の全てのマッチ箇所を Collection として返す
Public Function re_findall(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Collection
    Dim regex As Object, matches As Object, matchItem As Variant
    Dim results As New Collection
    Set regex = getRegex(pattern, ignoreCase, True)
    Set matches = regex.Execute(text)
    For Each matchItem In matches
        results.Add matchItem.Value
    Next matchItem
    Set re_findall = results
End Function

' re.sub: 対象文字列内のパターンにマッチする部分を置換する
Public Function re_sub(pattern As String, replacement As String, text As String, Optional ignoreCase As Boolean = True) As String
    Dim regex As Object
    Set regex = getRegex(pattern, ignoreCase, True)
    re_sub = regex.Replace(text, replacement)
End Function

' re.split: パターンにより文字列を分割し、各部分を Collection として返す
Public Function re_split(pattern As String, text As String, Optional ignoreCase As Boolean = True) As Collection
    Dim regex As Object, matches As Object
    Dim result As New Collection
    Dim lastEnd As Long, currentPos As Long
    Dim i As Long

    Set regex = getRegex(pattern, ignoreCase, True)
    Set matches = regex.Execute(text)

    lastEnd = 1
    For i = 0 To matches.Count - 1
        currentPos = matches(i).FirstIndex + 1
        result.Add Mid(text, lastEnd, currentPos - lastEnd)
        lastEnd = currentPos + matches(i).Length
    Next i
    result.Add Mid(text, lastEnd)
    Set re_split = result
End Function

' re_group: マッチオブジェクトから、指定した番号のキャプチャグループを返す
' groupNumber = 0 の場合は、マッチ全体を返します
Public Function re_group(matchObj As Object, Optional groupNumber As Integer = 0) As String
    Dim subMatches As Variant
    If matchObj Is Nothing Then
        re_group = ""
        Exit Function
    End If

    If groupNumber = 0 Then
        re_group = matchObj.Value
    Else
        subMatches = matchObj.SubMatches
        ' VBScriptのSubMatchesは0から始まる配列
        If groupNumber - 1 <= UBound(subMatches) Then
            re_group = subMatches(groupNumber - 1)
        Else
            re_group = ""
        End If
    End If
End Function

' re_groups: マッチオブジェクトから、すべてのキャプチャグループを Collection として返す
Public Function re_groups(matchObj As Object) As Collection
    Dim groups As New Collection
    Dim subMatches As Variant
    Dim i As Integer
    If matchObj Is Nothing Then
        Set re_groups = groups
        Exit Function
    End If

    subMatches = matchObj.SubMatches
    For i = LBound(subMatches) To UBound(subMatches)
        groups.Add subMatches(i)
    Next i
    Set re_groups = groups
End Function
