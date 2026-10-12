Attribute VB_Name = "mdlAT_Text"
Option Explicit


' @用途: Null・Empty・エラーを代替文字列にする
' @例: name = AT_TextOr(data(1, 1), "未入力")
Public Function AT_TextOr(ByVal value As Variant, Optional ByVal fallback As String = "") As String
    If IsError(value) Then AT_TextOr = fallback: Exit Function
    If IsNull(value) Then AT_TextOr = fallback: Exit Function
    If IsEmpty(value) Then AT_TextOr = fallback: Exit Function
    AT_TextOr = CStr(value)
End Function

' @用途: 全角スペース・タブ・連続スペースを整える
' @例: cleaned = AT_NormalizeSpaces("  A" & vbTab & "B  ")
Public Function AT_NormalizeSpaces(ByVal text As String) As String
    text = Replace(Replace(text, ChrW$(&H3000), " "), vbTab, " ")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    AT_NormalizeSpaces = Trim$(text)
End Function

' @用途: 改行コードを統一する。既定はセル内改行
' @例: text = AT_NormalizeNewlines(text, vbCrLf)
Public Function AT_NormalizeNewlines(ByVal text As String, Optional ByVal newline As String = vbLf) As String
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    AT_NormalizeNewlines = Replace(text, vbLf, newline)
End Function

' @用途: 大文字小文字を無視して部分一致を確認する
' @例: If AT_Contains(fileName, "売上") Then Debug.Print fileName
Public Function AT_Contains(ByVal text As String, ByVal part As String, Optional ByVal matchCase As Boolean = False) As Boolean
    Dim compare As Long
    If matchCase Then compare = vbBinaryCompare Else compare = vbTextCompare
    AT_Contains = InStr(1, text, part, compare) > 0
End Function

' @用途: 2つの区切り文字に挟まれた最初の文字列を取得する
' @例: code = AT_Between("品名[ABC]", "[", "]")
Public Function AT_Between(ByVal text As String, ByVal opening As String, ByVal closing As String) As String
    Dim first As Long, last As Long
    If Len(opening) = 0 Or Len(closing) = 0 Then Err.Raise 5, "AT_Between", "Delimiters must not be empty."
    first = InStr(1, text, opening, vbBinaryCompare)
    If first = 0 Then Exit Function
    first = first + Len(opening)
    last = InStr(first, text, closing, vbBinaryCompare)
    If last = 0 Then Exit Function
    AT_Between = Mid$(text, first, last - first)
End Function

' @用途: 文字列を左側から埋めて指定桁数にする
' @例: code = AT_PadLeft("7", 4, "0")
Public Function AT_PadLeft(ByVal text As String, ByVal width As Long, Optional ByVal pad As String = "0") As String
    If width < 0 Or Len(pad) <> 1 Then Err.Raise 5, "AT_PadLeft", "Invalid width or padding."
    If Len(text) >= width Then AT_PadLeft = text Else AT_PadLeft = String$(width - Len(text), pad) & text
End Function

' @用途: コードと名前を指定した区切り文字で結合する
' @例: label = AT_CombineCodeName("001", "サンプル", " / ")
Public Function AT_CombineCodeName(ByVal code As String, ByVal name As String, Optional ByVal separator As String = " / ") As String
    If Len(separator) = 0 Then Err.Raise 5, "AT_CombineCodeName", "Separator must not be empty."
    AT_CombineCodeName = code & separator & name
End Function

' @用途: 結合した表示文字列からコードを取り出す。区切りがなければ全体を返す
' @例: code = AT_CodeFromLabel("001 / サンプル")
Public Function AT_CodeFromLabel(ByVal label As String, Optional ByVal separator As String = " / ") As String
    Dim position As Long
    If Len(separator) = 0 Then Err.Raise 5, "AT_CodeFromLabel", "Separator must not be empty."
    position = InStr(1, label, separator, vbBinaryCompare)
    If position = 0 Then AT_CodeFromLabel = label Else AT_CodeFromLabel = Left$(label, position - 1)
End Function

' @用途: 結合した表示文字列から名前を取り出す
' @例: name = AT_NameFromLabel("001 / サンプル")
Public Function AT_NameFromLabel(ByVal label As String, Optional ByVal separator As String = " / ") As String
    Dim position As Long
    If Len(separator) = 0 Then Err.Raise 5, "AT_NameFromLabel", "Separator must not be empty."
    position = InStr(1, label, separator, vbBinaryCompare)
    If position > 0 Then AT_NameFromLabel = Mid$(label, position + Len(separator))
End Function

' @用途: 複数文字の区切りにも対応して文字列を分割する。0始まり
' @例: parts = AT_SplitText("A||B||", "||")
Public Function AT_SplitText(ByVal text As String, ByVal separator As String, Optional ByVal compare As Long = vbBinaryCompare) As Variant
    If Len(separator) = 0 Then AT_SplitText = Array(text) Else AT_SplitText = Split(text, separator, -1, compare)
End Function

' @用途: 文字列を置換する。検索文字が空ならそのまま返す
' @例: text = AT_ReplaceText(text, "旧", "新")
Public Function AT_ReplaceText(ByVal text As String, ByVal oldText As String, ByVal newText As String, Optional ByVal compare As Long = vbBinaryCompare) As String
    If Len(oldText) = 0 Then AT_ReplaceText = text Else AT_ReplaceText = Replace(text, oldText, newText, 1, -1, compare)
End Function

' @用途: HTMLのbrタグをセル内改行へ変換する
' @例: text = AT_BrToNewline("A<BR>B<br/>C")
Public Function AT_BrToNewline(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True: regex.IgnoreCase = True: regex.Pattern = "<br\s*/?>"
    AT_BrToNewline = regex.Replace(text, vbLf)
End Function

' @用途: 数値に変換できるか確認し、変換結果を受け取る
' @例: If AT_TryNumber(ws.Range("A1").Value2, amount) Then Debug.Print amount
Public Function AT_TryNumber(ByVal value As Variant, ByRef number As Double) As Boolean
    number = 0
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    If VarType(value) = vbBoolean Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    On Error GoTo Failed
    number = CDbl(value): AT_TryNumber = True
Failed:
End Function

' @用途: Excelの列番号をA・AA・XFD形式にする
' @例: Debug.Print AT_ColumnLetter(703)
Public Function AT_ColumnLetter(ByVal column As Long) As String
    Dim result As String
    If column < 1 Or column > 16384 Then Err.Raise 5, "AT_ColumnLetter", "Column must be 1..16384."
    Do While column > 0
        column = column - 1
        result = Chr$(65 + column Mod 26) & result
        column = column \ 26
    Loop
    AT_ColumnLetter = result
End Function

' @用途: A・AA・XFD形式の列名を列番号にする
' @例: Debug.Print AT_ColumnNumber("AAA")
Public Function AT_ColumnNumber(ByVal letters As String) As Long
    Dim i As Long, digit As Long, result As Long
    letters = UCase$(Trim$(letters))
    If Len(letters) < 1 Or Len(letters) > 3 Then Err.Raise 5, "AT_ColumnNumber", "Invalid column name."
    For i = 1 To Len(letters)
        digit = Asc(Mid$(letters, i, 1)) - 64
        If digit < 1 Or digit > 26 Then Err.Raise 5, "AT_ColumnNumber", "Invalid column name."
        result = result * 26 + digit
    Next i
    If result > 16384 Then Err.Raise 5, "AT_ColumnNumber", "Column exceeds XFD."
    AT_ColumnNumber = result
End Function

' @用途: CSVの1フィールドをエスケープする
' @例: Debug.Print AT_CsvField("A,B")
Public Function AT_CsvField(ByVal text As String, Optional ByVal delimiter As String = ",") As String
    If Len(delimiter) <> 1 Then Err.Raise 5, "AT_CsvField", "Delimiter must be one character."
    If InStr(text, delimiter) > 0 Or InStr(text, """") > 0 Or InStr(text, vbCr) > 0 Or InStr(text, vbLf) > 0 Then
        AT_CsvField = """" & Replace(text, """", """""") & """"
    Else
        AT_CsvField = text
    End If
End Function
