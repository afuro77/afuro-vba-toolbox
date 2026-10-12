Attribute VB_Name = "mdlAT_Range"
Option Explicit


' @用途: 値・数式が入った最終行を取得する。空なら0
' @例: last = AT_LastRow(ws, "A")
Public Function AT_LastRow(ByVal ws As Worksheet, Optional ByVal column As Variant = 0) As Long
    Dim scope As Range, found As Range
    If VarType(column) = vbString Then
        Set scope = ws.Columns(CStr(column))
    ElseIf CLng(column) = 0 Then
        Set scope = ws.Cells
    Else
        Set scope = ws.Columns(CLng(column))
    End If
    Set found = scope.Find(What:="*", After:=scope.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not found Is Nothing Then AT_LastRow = found.Row
End Function

' @用途: 値・数式が入った最終列を取得する。空なら0
' @例: last = AT_LastColumn(ws, 1)
Public Function AT_LastColumn(ByVal ws As Worksheet, Optional ByVal row As Long = 0) As Long
    Dim scope As Range, found As Range
    If row = 0 Then Set scope = ws.Cells Else Set scope = ws.Rows(row)
    Set found = scope.Find(What:="*", After:=scope.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not found Is Nothing Then AT_LastColumn = found.Column
End Function

' @用途: 指定した開始セルから最終行・列までの範囲を取得する。空ならNothing
' @例: Set data = AT_DataRange(ws, 2, 1)
Public Function AT_DataRange(ByVal ws As Worksheet, Optional ByVal firstRow As Long = 1, Optional ByVal firstColumn As Long = 1) As Range
    Dim lastRow As Long, lastColumn As Long
    If firstRow < 1 Or firstColumn < 1 Then Err.Raise 5, "AT_DataRange", "Invalid starting cell."
    lastRow = AT_LastRow(ws): lastColumn = AT_LastColumn(ws)
    If lastRow < firstRow Or lastColumn < firstColumn Then Exit Function
    Set AT_DataRange = ws.Range(ws.Cells(firstRow, firstColumn), ws.Cells(lastRow, lastColumn))
End Function

' @用途: 文字列を検索する。*や?も普通の文字として扱う
' @例: Set cell = AT_FindCell(ws.Range("A1:Z20"), "担当者")
Public Function AT_FindCell(ByVal scope As Range, ByVal text As String, Optional ByVal exact As Boolean = True, Optional ByVal matchCase As Boolean = False) As Range
    Dim key As String, lookAt As Long
    key = Replace(Replace(Replace(text, "~", "~~"), "*", "~*"), "?", "~?")
    If exact Then lookAt = xlWhole Else lookAt = xlPart
    Set AT_FindCell = scope.Find(What:=key, After:=scope.Cells(scope.Rows.Count, scope.Columns.Count), _
        LookIn:=xlValues, LookAt:=lookAt, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=matchCase, SearchFormat:=False)
End Function

' @用途: 見出し名から列番号を取得する。見つからなければ0
' @例: col = AT_HeaderColumn(ws, "金額", 1)
Public Function AT_HeaderColumn(ByVal ws As Worksheet, ByVal name As String, Optional ByVal headerRow As Long = 1) As Long
    Dim found As Range
    Set found = AT_FindCell(ws.Rows(headerRow), name)
    If Not found Is Nothing Then AT_HeaderColumn = found.Column
End Function

' @用途: 値だけを転記する。書式・数式はコピーしない
' @例: AT_CopyValues ws.Range("A2:D20"), dst.Range("B3")
Public Sub AT_CopyValues(ByVal source As Range, ByVal destination As Range)
    If source.Areas.Count <> 1 Then Err.Raise 5, "AT_CopyValues", "Source must be contiguous."
    destination.Cells(1, 1).Resize(source.Rows.Count, source.Columns.Count).Value2 = source.Value2
End Sub

' @用途: セル範囲を2次元配列にする。1セルでも配列を返す
' @例: data = AT_ReadMatrix(ws.Range("A1:D20"))
Public Function AT_ReadMatrix(ByVal source As Range) As Variant
    Dim oneCell(1 To 1, 1 To 1) As Variant
    If source.Areas.Count <> 1 Then Err.Raise 5, "AT_ReadMatrix", "Source must be contiguous."
    If source.CountLarge = 1 Then
        oneCell(1, 1) = source.Value2: AT_ReadMatrix = oneCell
    Else
        AT_ReadMatrix = source.Value2
    End If
End Function

' @用途: 2次元配列をまとめてセルに書く
' @例: AT_WriteMatrix dst.Range("A2"), data
Public Sub AT_WriteMatrix(ByVal destination As Range, ByVal data As Variant)
    Dim rows As Long, columns As Long
    rows = UBound(data, 1) - LBound(data, 1) + 1
    columns = UBound(data, 2) - LBound(data, 2) + 1
    destination.Cells(1, 1).Resize(rows, columns).Value2 = data
End Sub

' @用途: Nothingを含む2つの範囲を結合する
' @例: Set allCells = AT_UnionRange(allCells, ws.Range("C3"))
Public Function AT_UnionRange(ByVal left As Range, ByVal right As Range) As Range
    If left Is Nothing Then
        Set AT_UnionRange = right
    ElseIf right Is Nothing Then
        Set AT_UnionRange = left
    Else
        If Not left.Parent Is right.Parent Then Err.Raise 5, "AT_UnionRange", "Ranges belong to different sheets."
        Set AT_UnionRange = Application.Union(left, right)
    End If
End Function

' @用途: 指定範囲の完全な空行を下から削除する。数式のある行は残す
' @例: deleted = AT_RemoveEmptyRows(ws.Range("A2:H100"))
Public Function AT_RemoveEmptyRows(ByVal scope As Range) As Long
    Dim row As Long
    If scope.Areas.Count <> 1 Then Err.Raise 5, "AT_RemoveEmptyRows", "Range must be contiguous."
    For row = scope.Rows.Count To 1 Step -1
        If Application.CountA(scope.Rows(row)) = 0 Then
            scope.Rows(row).EntireRow.Delete
            AT_RemoveEmptyRows = AT_RemoveEmptyRows + 1
        End If
    Next row
End Function

' @用途: テンプレート行を挿入先へコピーする。値・数式・書式・行高を保持する
' @例: AT_InsertTemplateRow ws, 4, 20
Public Sub AT_InsertTemplateRow(ByVal ws As Worksheet, ByVal templateRow As Long, ByVal insertRow As Long)
    Dim actualTemplate As Long
    If templateRow < 1 Or insertRow < 1 Or templateRow > ws.Rows.Count Or insertRow > ws.Rows.Count Then Err.Raise 5, "AT_InsertTemplateRow", "Invalid row."
    If templateRow = ws.Rows.Count And insertRow <= templateRow Then Err.Raise 5, "AT_InsertTemplateRow", "The template would move outside the worksheet."
    ws.Rows(insertRow).Insert Shift:=xlDown
    actualTemplate = templateRow
    If insertRow <= templateRow Then actualTemplate = templateRow + 1
    ws.Rows(actualTemplate).Copy Destination:=ws.Rows(insertRow)
    ws.Rows(insertRow).RowHeight = ws.Rows(actualTemplate).RowHeight
    Application.CutCopyMode = False
End Sub

' @用途: 範囲を1列で昇順または降順に並べ替える
' @例: AT_SortRange ws.Range("A1:D100"), 2, True, True
Public Sub AT_SortRange(ByVal scope As Range, ByVal keyColumn As Long, Optional ByVal hasHeader As Boolean = True, Optional ByVal ascending As Boolean = True)
    Dim direction As Long, header As Long
    If keyColumn < 1 Or keyColumn > scope.Columns.Count Then Err.Raise 5, "AT_SortRange", "Key column is relative to the range."
    If ascending Then direction = xlAscending Else direction = xlDescending
    If hasHeader Then header = xlYes Else header = xlNo
    scope.Sort Key1:=scope.Columns(keyColumn), Order1:=direction, Header:=header, Orientation:=xlTopToBottom
End Sub
