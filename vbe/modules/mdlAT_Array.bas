Attribute VB_Name = "mdlAT_Array"
Option Explicit


' @用途: 1次元配列の要素数を取得する。空・未初期化なら0
' @例: count = AT_ArrayLength(Array("A", "B"))
Public Function AT_ArrayLength(ByVal values As Variant) As Long
    If Not IsArray(values) Then Err.Raise 5, "AT_ArrayLength", "Array is required."
    On Error GoTo EmptyArray
    AT_ArrayLength = UBound(values) - LBound(values) + 1
EmptyArray:
End Function

' @用途: Collectionを0始まりの配列にする
' @例: values = AT_CollectionToArray(items)
Public Function AT_CollectionToArray(ByVal items As Collection) As Variant
    Dim result() As Variant, i As Long
    If items.Count = 0 Then AT_CollectionToArray = Array(): Exit Function
    ReDim result(0 To items.Count - 1)
    For i = 1 To items.Count
        If IsObject(items(i)) Then Set result(i - 1) = items(i) Else result(i - 1) = items(i)
    Next i
    AT_CollectionToArray = result
End Function

' @用途: 1次元配列から重複を除く。文字列として比較する
' @例: unique = AT_UniqueValues(Array("A", "a", "B"))
Public Function AT_UniqueValues(ByVal values As Variant, Optional ByVal matchCase As Boolean = False) As Variant
    Dim dictionary As Object, item As Variant, key As String
    Set dictionary = CreateObject("Scripting.Dictionary")
    If matchCase Then dictionary.CompareMode = vbBinaryCompare Else dictionary.CompareMode = vbTextCompare
    For Each item In values
        If IsError(item) Or IsNull(item) Or IsObject(item) Or IsArray(item) Then Err.Raise 5, "AT_UniqueValues", "Use scalar non-error values."
        key = CStr(item)
        If Not dictionary.Exists(key) Then dictionary.Add key, item
    Next item
    AT_UniqueValues = dictionary.Items
End Function

' @用途: 1次元配列から空文字・Empty・Nullだけを除く。0は残す
' @例: values = AT_FilterNonEmpty(Array("", 0, "A"))
Public Function AT_FilterNonEmpty(ByVal values As Variant) As Variant
    Dim items As New Collection, item As Variant
    For Each item In values
        If IsError(item) Then
            items.Add item
        ElseIf Not IsNull(item) And Not IsEmpty(item) Then
            If CStr(item) <> "" Then items.Add item
        End If
    Next item
    AT_FilterNonEmpty = AT_CollectionToArray(items)
End Function

' @用途: 2次元配列を転置する。出力は1始まり
' @例: transposed = AT_TransposeMatrix(data)
Public Function AT_TransposeMatrix(ByVal data As Variant) As Variant
    Dim result() As Variant, row As Long, column As Long, firstRow As Long, firstColumn As Long
    firstRow = LBound(data, 1): firstColumn = LBound(data, 2)
    ReDim result(1 To UBound(data, 2) - firstColumn + 1, 1 To UBound(data, 1) - firstRow + 1)
    For row = firstRow To UBound(data, 1)
        For column = firstColumn To UBound(data, 2)
            result(column - firstColumn + 1, row - firstRow + 1) = data(row, column)
        Next column
    Next row
    AT_TransposeMatrix = result
End Function

' @用途: 2次元配列から1列を取り出す。出力は1始まり
' @例: codes = AT_MatrixColumn(data, 1)
Public Function AT_MatrixColumn(ByVal data As Variant, ByVal column As Long) As Variant
    Dim result() As Variant, row As Long, first As Long
    If column < LBound(data, 2) Or column > UBound(data, 2) Then Err.Raise 5, "AT_MatrixColumn", "Invalid column."
    first = LBound(data, 1)
    ReDim result(1 To UBound(data, 1) - first + 1)
    For row = first To UBound(data, 1)
        result(row - first + 1) = data(row, column)
    Next row
    AT_MatrixColumn = result
End Function

' @用途: 2次元配列の指定列がキーと一致する行だけを返す。該当なしはEmpty
' @例: selected = AT_FilterMatrix(data, 1, "A001")
Public Function AT_FilterMatrix(ByVal data As Variant, ByVal column As Long, ByVal key As String, Optional ByVal matchCase As Boolean = False) As Variant
    Dim selected As New Collection, row As Long, col As Long, result() As Variant, i As Long, compare As Long
    If column < LBound(data, 2) Or column > UBound(data, 2) Then Err.Raise 5, "AT_FilterMatrix", "Invalid column."
    If matchCase Then compare = vbBinaryCompare Else compare = vbTextCompare
    For row = LBound(data, 1) To UBound(data, 1)
        If Not IsError(data(row, column)) And Not IsNull(data(row, column)) Then
            If StrComp(CStr(data(row, column)), key, compare) = 0 Then selected.Add row
        End If
    Next row
    If selected.Count = 0 Then Exit Function
    ReDim result(1 To selected.Count, 1 To UBound(data, 2) - LBound(data, 2) + 1)
    For i = 1 To selected.Count
        For col = LBound(data, 2) To UBound(data, 2)
            result(i, col - LBound(data, 2) + 1) = data(selected(i), col)
        Next col
    Next i
    AT_FilterMatrix = result
End Function
