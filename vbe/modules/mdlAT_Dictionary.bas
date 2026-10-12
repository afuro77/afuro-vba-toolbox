Attribute VB_Name = "mdlAT_Dictionary"
Option Explicit


' @用途: 参照設定不要の辞書を作る。既定は大文字小文字を区別しない
' @例: Set index = AT_NewDictionary()
Public Function AT_NewDictionary(Optional ByVal matchCase As Boolean = False) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    If matchCase Then result.CompareMode = vbBinaryCompare Else result.CompareMode = vbTextCompare
    Set AT_NewDictionary = result
End Function

' @用途: 1次元配列の値ごとの件数を集計する。文字列として比較する
' @例: Set counts = AT_CountValues(Array("A", "B", "A"))
Public Function AT_CountValues(ByVal values As Variant, Optional ByVal matchCase As Boolean = False) As Object
    Dim result As Object, item As Variant, key As String
    Set result = AT_NewDictionary(matchCase)
    For Each item In values
        If IsError(item) Or IsNull(item) Or IsObject(item) Or IsArray(item) Then Err.Raise 5, "AT_CountValues", "Use scalar non-error values."
        key = CStr(item)
        If result.Exists(key) Then result(key) = CLng(result(key)) + 1 Else result.Add key, 1&
    Next item
    Set AT_CountValues = result
End Function

' @用途: 2次元配列のキーから行番号を引ける辞書を作る。重複は既定でエラー
' @例: Set index = AT_IndexRows(data, 1, "ERROR")
Public Function AT_IndexRows(ByVal data As Variant, ByVal keyColumn As Long, Optional ByVal duplicates As String = "ERROR") As Object
    Dim result As Object, row As Long, key As String, value As Variant
    duplicates = UCase$(duplicates)
    If duplicates <> "ERROR" And duplicates <> "FIRST" And duplicates <> "LAST" Then Err.Raise 5, "AT_IndexRows", "Use ERROR, FIRST or LAST."
    If keyColumn < LBound(data, 2) Or keyColumn > UBound(data, 2) Then Err.Raise 5, "AT_IndexRows", "Invalid key column."
    Set result = AT_NewDictionary()
    For row = LBound(data, 1) To UBound(data, 1)
        value = data(row, keyColumn)
        If IsError(value) Or IsNull(value) Then Err.Raise 5, "AT_IndexRows", "Invalid key at row " & row
        key = CStr(value)
        If result.Exists(key) Then
            Select Case duplicates
                Case "ERROR": Err.Raise 5, "AT_IndexRows", "Duplicate key: " & key
                Case "LAST": result(key) = row
            End Select
        Else
            result.Add key, row
        End If
    Next row
    Set AT_IndexRows = result
End Function

' @用途: キー別に数値列を合計する。空値・文字列はエラーとして知らせる
' @例: Set totals = AT_GroupSum(data, 1, 3)
Public Function AT_GroupSum(ByVal data As Variant, ByVal keyColumn As Long, ByVal valueColumn As Long) As Object
    Dim result As Object, row As Long, key As String, value As Variant
    If keyColumn < LBound(data, 2) Or keyColumn > UBound(data, 2) Then Err.Raise 5, "AT_GroupSum", "Invalid key column."
    If valueColumn < LBound(data, 2) Or valueColumn > UBound(data, 2) Then Err.Raise 5, "AT_GroupSum", "Invalid value column."
    Set result = AT_NewDictionary()
    For row = LBound(data, 1) To UBound(data, 1)
        value = data(row, keyColumn)
        If IsError(value) Or IsNull(value) Then Err.Raise 5, "AT_GroupSum", "Invalid key at row " & row
        key = CStr(value): value = data(row, valueColumn)
        If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Err.Raise 5, "AT_GroupSum", "Missing number at row " & row
        If VarType(value) = vbBoolean Or Not IsNumeric(value) Then Err.Raise 5, "AT_GroupSum", "Invalid number at row " & row
        If result.Exists(key) Then result(key) = CDbl(result(key)) + CDbl(value) Else result.Add key, CDbl(value)
    Next row
    Set AT_GroupSum = result
End Function

' @用途: 辞書をキー・値の2列配列へ変換する。空ならEmpty
' @例: data = AT_DictionaryMatrix(totals)
Public Function AT_DictionaryMatrix(ByVal dictionary As Object) As Variant
    Dim result() As Variant, key As Variant, row As Long
    If dictionary.Count = 0 Then Exit Function
    ReDim result(1 To dictionary.Count, 1 To 2)
    For Each key In dictionary.Keys
        row = row + 1: result(row, 1) = key: result(row, 2) = dictionary(key)
    Next key
    AT_DictionaryMatrix = result
End Function
