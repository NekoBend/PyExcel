' クラスモジュール: Regex
Option Explicit

Private reg As Object  ' VBScript.RegExp オブジェクト
Private mPattern As String
Private mFlags As Long

' フラグの定数
Public Const IGNORECASE As Long = 1
Public Const MULTILINE As Long = 2
Public Const DOTALL As Long = 4

' 初期化メソッド
Public Sub Init(pattern As String, Optional flags As Long = 0)
    mPattern = pattern
    mFlags = flags
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = AdjustPatternForFlags(pattern, flags)
    reg.IgnoreCase = ((flags And IGNORECASE) <> 0)
    reg.Global = True  ' 全ての一致を取得するため
End Sub

' DOTALL フラグ対応のため、文字クラス外の "." を "[\s\S]" に置換（単純な実装）
Private Function AdjustPatternForFlags(pattern As String, flags As Long) As String
    Dim newPattern As String
    newPattern = pattern
    If (flags And DOTALL) <> 0 Then
        Dim i As Long, c As String, inCharClass As Boolean, prevChar As String
        Dim result As String
        inCharClass = False
        result = ""
        prevChar = ""
        For i = 1 To Len(newPattern)
            c = Mid(newPattern, i, 1)
            If c = "[" And prevChar <> "\" Then
                inCharClass = True
            ElseIf c = "]" And inCharClass Then
                inCharClass = False
            End If
            If c = "." And Not inCharClass And prevChar <> "\" Then
                result = result & "[\\s\\S]"
            Else
                result = result & c
            End If
            prevChar = c
        Next i
        AdjustPatternForFlags = result
    Else
        AdjustPatternForFlags = pattern
    End If
End Function

' Search: 文字列内の任意の位置で最初の一致を返す
Public Function Search(text As String) As Object
    Dim matches As Object
    Set matches = reg.Execute(text)
    If matches.Count > 0 Then
        Set Search = matches.Item(0)
    Else
        Set Search = Nothing
    End If
End Function

' Match: 文字列の先頭で一致する場合のみ返す
Public Function Match(text As String) As Object
    Dim m As Object
    Set m = Me.Search(text)
    If Not m Is Nothing Then
        If m.FirstIndex = 0 Then
            Set Match = m
        Else
            Set Match = Nothing
        End If
    Else
        Set Match = Nothing
    End If
End Function

' FullMatch: 文字列全体がパターンに一致する場合のみ返す
Public Function FullMatch(text As String) As Object
    Dim m As Object
    Dim fullPattern As String
    fullPattern = "^" & mPattern & "$"
    Dim tempReg As Object
    Set tempReg = CreateObject("VBScript.RegExp")
    tempReg.Pattern = AdjustPatternForFlags(fullPattern, mFlags)
    tempReg.IgnoreCase = reg.IgnoreCase
    tempReg.Global = False
    Set m = tempReg.Execute(text)
    If m.Count > 0 Then
        Set FullMatch = m.Item(0)
    Else
        Set FullMatch = Nothing
    End If
End Function

' FindAll: すべての一致を Collection として返す
Public Function FindAll(text As String) As Collection
    Dim coll As New Collection
    Dim matches As Object, m As Object
    Set matches = reg.Execute(text)
    For Each m In matches
        coll.Add m
    Next m
    Set FindAll = coll
End Function

' FindIter: FindAll と同等（イテレータ的に Collection を返す）
Public Function FindIter(text As String) As Collection
    Set FindIter = Me.FindAll(text)
End Function

' Split: パターンにより文字列を分割して配列として返す
Public Function Split(text As String) As Variant
    Dim matches As Object
    Dim result() As String
    Dim i As Long, startPos As Long
    Set matches = reg.Execute(text)
    startPos = 1
    ReDim result(0)
    i = 0
    Dim m As Object
    For Each m In matches
        ReDim Preserve result(i)
        result(i) = Mid(text, startPos, m.FirstIndex - startPos + 1)
        i = i + 1
        startPos = m.FirstIndex + m.Length + 1
    Next m
    ReDim Preserve result(i)
    result(i) = Mid(text, startPos)
    Split = result
End Function

' Sub: 一致した部分を replacement で置換
Public Function Sub(replacement As String, text As String) As String
    Sub = reg.Replace(text, replacement)
End Function

' SubN: 置換後の文字列と置換回数を配列で返す（配列(0)=新しい文字列、(1)=置換回数）
Public Function SubN(replacement As String, text As String) As Variant
    Dim newText As String
    Dim matches As Object
    Set matches = reg.Execute(text)
    newText = reg.Replace(text, replacement)
    Dim result(1) As Variant
    result(0) = newText
    result(1) = matches.Count
    SubN = result
End Function