Attribute VB_Name = "mdlAT_All"
Option Explicit

Public Type AT_ExcelState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As Long
    StatusBar As Variant
    Captured As Boolean
End Type


' ===== mdlAT_Array =====
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

' ===== mdlAT_Csv =====
' @用途: 2次元配列をCSV文字列へ変換する。引用符・カンマ・改行に対応
' @例: csv = AT_MatrixToCsv(data)
Public Function AT_MatrixToCsv(ByVal data As Variant, Optional ByVal delimiter As String = ",") As String
    Dim row As Long, column As Long, field As Variant, line As String, result As String
    If Len(delimiter) <> 1 Or delimiter = """" Or delimiter = vbCr Or delimiter = vbLf Then Err.Raise 5, "AT_MatrixToCsv", "Invalid delimiter."
    For row = LBound(data, 1) To UBound(data, 1)
        line = ""
        For column = LBound(data, 2) To UBound(data, 2)
            field = data(row, column)
            If IsError(field) Or IsObject(field) Or IsArray(field) Then Err.Raise 5, "AT_MatrixToCsv", "Invalid field at row " & row
            If column > LBound(data, 2) Then line = line & delimiter
            line = line & AT_CsvField(AT_TextOr(field), delimiter)
        Next column
        If row > LBound(data, 1) Then result = result & vbCrLf
        result = result & line
    Next row
    AT_MatrixToCsv = result
End Function

' @用途: CSV文字列を2次元配列へ変換する。空入力ならEmpty
' @例: data = AT_ParseCsv(text)
Public Function AT_ParseCsv(ByVal text As String, Optional ByVal delimiter As String = ",") As Variant
    Dim records As New Collection, fields As New Collection, result() As Variant
    Dim value As String, quoted As Boolean, closed As Boolean, fieldStarted As Boolean
    Dim character As String, i As Long, maxColumns As Long, row As Long, column As Long, record As Collection
    If Len(delimiter) <> 1 Or delimiter = """" Or delimiter = vbCr Or delimiter = vbLf Then Err.Raise 5, "AT_ParseCsv", "Invalid delimiter."
    If Len(text) = 0 Then Exit Function
    ' Strip a decoded Unicode BOM, if present.
    If Left$(text, 1) = ChrW$(&HFEFF) Then text = Mid$(text, 2)
    If Len(text) = 0 Then Exit Function
    i = 1
    Do While i <= Len(text)
        character = Mid$(text, i, 1)
        If quoted Then
            If character = """" Then
                If Mid$(text, i + 1, 1) = """" Then
                    value = value & """": i = i + 1
                Else
                    quoted = False: closed = True
                End If
            Else
                value = value & character
            End If
        ElseIf character = delimiter Then
            fields.Add value: value = "": closed = False: fieldStarted = False
        ElseIf character = vbCr Or character = vbLf Then
            fields.Add value: records.Add fields
            If fields.Count > maxColumns Then maxColumns = fields.Count
            Set fields = New Collection
            value = "": closed = False: fieldStarted = False
            If character = vbCr And Mid$(text, i + 1, 1) = vbLf Then i = i + 1
        ElseIf character = """" Then
            If fieldStarted Or closed Then Err.Raise 5, "AT_ParseCsv", "Unexpected quote at position " & i
            quoted = True: fieldStarted = True
        Else
            If closed Then Err.Raise 5, "AT_ParseCsv", "Unexpected text after closing quote at position " & i
            value = value & character: fieldStarted = True
        End If
        i = i + 1
    Loop
    If quoted Then Err.Raise 5, "AT_ParseCsv", "Unclosed quoted field."
    If fields.Count > 0 Or fieldStarted Or closed Or Len(value) > 0 Then
        fields.Add value: records.Add fields
        If fields.Count > maxColumns Then maxColumns = fields.Count
    End If
    ReDim result(1 To records.Count, 1 To maxColumns)
    For row = 1 To records.Count
        Set record = records(row)
        For column = 1 To maxColumns
            If column <= record.Count Then result(row, column) = record(column) Else result(row, column) = ""
        Next column
    Next row
    AT_ParseCsv = result
End Function

' @用途: 範囲の値をCSVとしてUTF-8で保存する。既定で上書きしない
' @例: AT_ExportRangeCsv ws.Range("A1:D20"), outputPath
Public Sub AT_ExportRangeCsv(ByVal source As Range, ByVal path As String, Optional ByVal overwrite As Boolean = False)
    AT_WriteUtf8 path, AT_MatrixToCsv(AT_ReadMatrix(source)), overwrite
End Sub

' ===== mdlAT_Date =====
' @用途: YYYY-MM-DDを厳密に日付へ変換する。不正な日付はFalse
' @例: If AT_TryIsoDate("2026-10-12", value) Then Debug.Print value
Public Function AT_TryIsoDate(ByVal text As String, ByRef value As Date) As Boolean
    Dim year As Long, month As Long, day As Long, candidate As Date
    value = 0
    If Not text Like "####-##-##" Or Len(text) <> 10 Then Exit Function
    year = CLng(Left$(text, 4)): month = CLng(Mid$(text, 6, 2)): day = CLng(Right$(text, 2))
    If year < 1000 Or month < 1 Or month > 12 Or day < 1 Or day > 31 Then Exit Function
    On Error GoTo InvalidDate
    candidate = DateSerial(year, month, day)
    If VBA.Year(candidate) = year And VBA.Month(candidate) = month And VBA.Day(candidate) = day Then
        value = candidate: AT_TryIsoDate = True
    End If
InvalidDate:
End Function

' @用途: 対象月の月初を取得する
' @例: first = AT_MonthStart(Date)
Public Function AT_MonthStart(ByVal value As Date) As Date
    AT_MonthStart = DateSerial(Year(value), Month(value), 1)
End Function

' @用途: 対象月の月末を取得する
' @例: last = AT_MonthEnd(Date)
Public Function AT_MonthEnd(ByVal value As Date) As Date
    AT_MonthEnd = DateSerial(Year(value), Month(value) + 1, 0)
End Function

' @用途: 対象週の月曜日を取得する。時刻は除く
' @例: monday = AT_WeekMonday(Date)
Public Function AT_WeekMonday(ByVal value As Date) As Date
    AT_WeekMonday = DateValue(value) - Weekday(value, vbMonday) + 1
End Function

' @用途: ファイル名に使える日時文字列を作る
' @例: stamp = AT_DateStamp(Now)
Public Function AT_DateStamp(ByVal value As Date) As String
    AT_DateStamp = Format$(value, "yyyymmdd_hhnnss")
End Function

' @用途: YYYYMM・YYYYMMDD・YYYYMMDDHHMMを表示用文字列へ変換する
' @例: Debug.Print AT_FormatCompactDate("202610121530")
Public Function AT_FormatCompactDate(ByVal text As String) As String
    Dim year As Long, month As Long, day As Long, candidate As Date, hour As Long, minute As Long
    If Len(text) <> 6 And Len(text) <> 8 And Len(text) <> 12 Then Err.Raise 5, "AT_FormatCompactDate", "Use 6, 8 or 12 digits."
    If text Like "*[!0-9]*" Then Err.Raise 5, "AT_FormatCompactDate", "Use digits only."
    year = CLng(Left$(text, 4)): month = CLng(Mid$(text, 5, 2)): day = 1
    If year < 1000 Or month < 1 Or month > 12 Then Err.Raise 5, "AT_FormatCompactDate", "Invalid year or month."
    If Len(text) >= 8 Then day = CLng(Mid$(text, 7, 2))
    candidate = DateSerial(year, month, day)
    If VBA.Month(candidate) <> month Or day < 1 Then Err.Raise 5, "AT_FormatCompactDate", "Invalid day."
    If Len(text) = 6 Then AT_FormatCompactDate = Format$(candidate, "yyyy/mm"): Exit Function
    If Len(text) = 8 Then AT_FormatCompactDate = Format$(candidate, "yyyy/mm/dd"): Exit Function
    hour = CLng(Mid$(text, 9, 2)): minute = CLng(Right$(text, 2))
    If hour > 23 Or minute > 59 Then Err.Raise 5, "AT_FormatCompactDate", "Invalid time."
    AT_FormatCompactDate = Format$(candidate + TimeSerial(hour, minute, 0), "yyyy/mm/dd hh:nn")
End Function

' ===== mdlAT_Dictionary =====
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

' ===== mdlAT_Excel =====
' @用途: Excelの高速化設定を開始し、元の設定を保存する
' @例: Dim state As AT_ExcelState: state = AT_PauseExcel()
Public Function AT_PauseExcel() As AT_ExcelState
    Dim state As AT_ExcelState, number As Long, message As String
    state.ScreenUpdating = Application.ScreenUpdating
    state.EnableEvents = Application.EnableEvents
    state.DisplayAlerts = Application.DisplayAlerts
    state.Calculation = Application.Calculation
    state.StatusBar = Application.StatusBar
    state.Captured = True
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    AT_PauseExcel = state
    Exit Function
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    AT_RestoreExcel state
    On Error GoTo 0
    Err.Raise number, "AT_PauseExcel", message
End Function

' @用途: 高速化前のExcel設定に戻す。エラー処理でも呼ぶ
' @例: AT_RestoreExcel state
Public Sub AT_RestoreExcel(ByRef state As AT_ExcelState)
    Dim number As Long, message As String
    If Not state.Captured Then Exit Sub
    ' Restore every property, even if a particular setting is unavailable.
    On Error Resume Next
    Application.Calculation = state.Calculation
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.DisplayAlerts = state.DisplayAlerts
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.EnableEvents = state.EnableEvents
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.ScreenUpdating = state.ScreenUpdating
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.StatusBar = state.StatusBar
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "AT_RestoreExcel", message
    state.Captured = False
End Sub

' @用途: ブック内にシートがあるか確認する
' @例: Debug.Print AT_HasSheet(ThisWorkbook, "集計")
Public Function AT_HasSheet(ByVal book As Workbook, ByVal name As String) As Boolean
    Dim sheet As Object
    On Error Resume Next
    Set sheet = book.Sheets(name)
    On Error GoTo 0
    AT_HasSheet = Not sheet Is Nothing
End Function

' @用途: ワークシートを取得する。必要なら作成する
' @例: Set ws = AT_GetSheet(ThisWorkbook, "集計", True)
Public Function AT_GetSheet(ByVal book As Workbook, ByVal name As String, Optional ByVal create As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = book.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing And create Then
        If AT_HasSheet(book, name) Then Err.Raise 5, "AT_GetSheet", "A non-worksheet already uses this name."
        Set ws = book.Worksheets.Add(After:=book.Sheets(book.Sheets.Count))
        On Error GoTo FailedName
        ws.Name = name
        On Error GoTo 0
    End If
    Set AT_GetSheet = ws
    Exit Function
FailedName:
    Dim number As Long, message As String, alerts As Boolean
    number = Err.Number: message = Err.Description
    alerts = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = alerts
    On Error GoTo 0
    Err.Raise number, "AT_GetSheet", message
End Function

' @用途: シート名に使えない文字を除き、31文字までにする
' @例: Debug.Print AT_SafeSheetName("売上/2026:10")
Public Function AT_SafeSheetName(ByVal value As String, Optional ByVal fallback As String = "Sheet") As String
    Dim mark As Variant, i As Long, result As String
    result = value
    For Each mark In Array(":", "\", "/", "?", "*", "[", "]")
        result = Replace(result, CStr(mark), "_")
    Next mark
    For i = 0 To 31
        result = Replace(result, Chr$(i), "")
    Next i
    result = Trim$(result)
    Do While Left$(result, 1) = "'"
        result = Mid$(result, 2)
    Loop
    Do While Right$(result, 1) = "'"
        result = Left$(result, Len(result) - 1)
    Loop
    If Len(result) = 0 Then
        If fallback <> "Sheet" Then result = AT_SafeSheetName(fallback, "Sheet") Else result = "Sheet"
    End If
    result = Left$(result, 31)
    Do While Right$(result, 1) = "'"
        result = Left$(result, Len(result) - 1)
    Loop
    ' History is reserved by some Excel workbook features.
    If StrComp(result, "History", vbTextCompare) = 0 Then result = "History_"
    AT_SafeSheetName = result
End Function

' @用途: 既存シートと重ならない名前を作る
' @例: newName = AT_UniqueSheetName(ThisWorkbook, "集計")
Public Function AT_UniqueSheetName(ByVal book As Workbook, ByVal requested As String) As String
    Dim base As String, candidate As String, suffix As String, i As Long
    base = AT_SafeSheetName(requested): candidate = base
    Do While AT_HasSheet(book, candidate)
        i = i + 1: suffix = "_" & CStr(i)
        candidate = Left$(base, 31 - Len(suffix)) & suffix
    Loop
    AT_UniqueSheetName = candidate
End Function

' @用途: シート名一覧をCollectionで取得する
' @例: Set names = AT_SheetNames(ThisWorkbook)
Public Function AT_SheetNames(ByVal book As Workbook) As Collection
    Dim result As New Collection, sheet As Object
    For Each sheet In book.Sheets
        result.Add sheet.Name
    Next sheet
    Set AT_SheetNames = result
End Function

' @用途: 名前またはフルパスで、開いているブックを探す
' @例: Set book = AT_OpenBookByName("sample.xlsx")
Public Function AT_OpenBookByName(ByVal nameOrPath As String) As Workbook
    Dim book As Workbook
    For Each book In Application.Workbooks
        If StrComp(book.Name, nameOrPath, vbTextCompare) = 0 Or StrComp(book.FullName, nameOrPath, vbTextCompare) = 0 Then
            Set AT_OpenBookByName = book
            Exit Function
        End If
    Next book
End Function

' @用途: 状態バーへ進捗を表示する。終了時は元の設定を復元する
' @例: AT_ShowProgress 3, 10, "処理中"
Public Sub AT_ShowProgress(ByVal current As Long, ByVal total As Long, Optional ByVal message As String = "処理中")
    If total <= 0 Or current < 0 Or current > total Then Err.Raise 5, "AT_ShowProgress", "Invalid progress."
    Application.StatusBar = message & " " & current & "/" & total & " (" & Format$(current / total, "0%") & ")"
End Sub

' ===== mdlAT_Files =====
' @用途: フルパスを結合する。後半は相対パスを指定する
' @例: path = AT_PathJoin("C:\sample", "data\input.csv")
Public Function AT_PathJoin(ByVal folder As String, ByVal child As String) As String
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Left$(child, 1) = "\" Or InStr(child, ":") > 0 Then Err.Raise 5, "AT_PathJoin", "Child must be a relative path."
    AT_PathJoin = fs.BuildPath(folder, child)
End Function

' @用途: ファイルの存在を確認する
' @例: If AT_FileExists(path) Then Debug.Print path
Public Function AT_FileExists(ByVal path As String) As Boolean
    AT_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' @用途: フォルダの存在を確認する
' @例: Debug.Print AT_FolderExists("C:\sample")
Public Function AT_FolderExists(ByVal path As String) As Boolean
    AT_FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(path)
End Function

' @用途: 親フォルダも含めてフォルダを作る。既存なら何もしない
' @例: AT_EnsureFolder "C:\sample\output\2026"
Public Sub AT_EnsureFolder(ByVal path As String)
    Dim fs As Object, parent As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    path = fs.GetAbsolutePathName(path)
    If fs.FolderExists(path) Then Exit Sub
    parent = fs.GetParentFolderName(path)
    If Len(parent) = 0 Or parent = path Then Err.Raise 76, "AT_EnsureFolder", "Root folder is unavailable."
    AT_EnsureFolder parent
    fs.CreateFolder path
End Sub

' @用途: ファイル名をフルパスから取得する
' @例: Debug.Print AT_FileName("C:\sample\input.csv")
Public Function AT_FileName(ByVal path As String) As String
    AT_FileName = CreateObject("Scripting.FileSystemObject").GetFileName(path)
End Function

' @用途: 拡張子なしのファイル名を取得する
' @例: Debug.Print AT_BaseName("C:\sample\input.csv")
Public Function AT_BaseName(ByVal path As String) As String
    AT_BaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(path)
End Function

' @用途: 拡張子を小文字で取得する。先頭のドットは含めない
' @例: Debug.Print AT_Extension("input.CSV")
Public Function AT_Extension(ByVal path As String) As String
    AT_Extension = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(path))
End Function

' @用途: ファイルの親フォルダを取得する
' @例: Debug.Print AT_ParentFolder("C:\sample\input.csv")
Public Function AT_ParentFolder(ByVal path As String) As String
    AT_ParentFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(path)
End Function

' @用途: 条件に一致するファイルをCollectionで取得する。再帰は任意
' @例: Set files = AT_ListFiles("C:\sample", "*.csv", True)
Public Function AT_ListFiles(ByVal folder As String, Optional ByVal pattern As String = "*", Optional ByVal recursive As Boolean = False) As Collection
    Dim result As New Collection, fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    AT_FilesCollect fs.GetFolder(folder), pattern, recursive, result
    Set AT_ListFiles = result
End Function

Private Sub AT_FilesCollect(ByVal folder As Object, ByVal pattern As String, ByVal recursive As Boolean, ByVal result As Collection)
    Dim file As Object, child As Object
    For Each file In folder.Files
        If LCase$(file.Name) Like LCase$(pattern) Then
            If Left$(file.Name, 2) <> "~$" Then result.Add file.Path
        End If
    Next file
    If recursive Then
        For Each child In folder.SubFolders
            If (child.Attributes And 1024) = 0 Then AT_FilesCollect child, pattern, True, result
        Next child
    End If
End Sub

' @用途: 既存ファイルと重ならないパスを作る。ファイル自体は作成しない
' @例: output = AT_UniqueFilePath("C:\sample\output.csv")
Public Function AT_UniqueFilePath(ByVal path As String) As String
    Dim fs As Object, stem As String, extension As String, candidate As String, i As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    stem = fs.BuildPath(fs.GetParentFolderName(path), fs.GetBaseName(path))
    extension = fs.GetExtensionName(path)
    If Len(extension) > 0 Then extension = "." & extension
    candidate = path
    Do While fs.FileExists(candidate) Or fs.FolderExists(candidate)
        i = i + 1: candidate = stem & "_" & i & extension
    Loop
    AT_UniqueFilePath = candidate
End Function

' @用途: ファイルをコピーする。既定では上書きしない
' @例: AT_CopyFile sourcePath, destinationPath
Public Sub AT_CopyFile(ByVal source As String, ByVal destination As String, Optional ByVal overwrite As Boolean = False)
    CreateObject("Scripting.FileSystemObject").CopyFile source, destination, overwrite
End Sub

' @用途: テキストファイルを指定文字コードで読む。既定はUTF-8
' @例: text = AT_ReadTextFile(path, "utf-8")
Public Function AT_ReadTextFile(ByVal path As String, Optional ByVal charset As String = "utf-8") As String
    Dim stream As Object, number As Long, message As String
    Set stream = CreateObject("ADODB.Stream")
    On Error GoTo Failed
    stream.Type = 2: stream.Charset = charset: stream.Open
    stream.LoadFromFile path
    AT_ReadTextFile = stream.ReadText(-1)
    stream.Close
    Exit Function
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    stream.Close
    On Error GoTo 0
    Err.Raise number, "AT_ReadTextFile", message
End Function

' @用途: UTF-8でテキストを書き出す。既定はBOMなし・上書き禁止
' @例: AT_WriteUtf8 path, text
Public Sub AT_WriteUtf8(ByVal path As String, ByVal text As String, Optional ByVal overwrite As Boolean = False, Optional ByVal includeBom As Boolean = False)
    Dim stream As Object, binary As Object, number As Long, message As String, saveMode As Long
    Set stream = CreateObject("ADODB.Stream")
    On Error GoTo Failed
    If overwrite Then saveMode = 2 Else saveMode = 1
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText text
    If includeBom Then
        stream.SaveToFile path, saveMode
    Else
        stream.Position = 0: stream.Type = 1
        Set binary = CreateObject("ADODB.Stream")
        binary.Type = 1: binary.Open
        If stream.Size >= 3 Then
            stream.Position = 3
            If stream.Size > 3 Then binary.Write stream.Read
        End If
        binary.SaveToFile path, saveMode
        binary.Close
    End If
    stream.Close
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    stream.Close
    If Not binary Is Nothing Then binary.Close
    On Error GoTo 0
    Err.Raise number, "AT_WriteUtf8", message
End Sub

' @用途: Windowsのファイル名に使えない文字を置換する
' @例: name = AT_SafeFileName("売上:2026/10.csv")
Public Function AT_SafeFileName(ByVal name As String) As String
    Dim mark As Variant, i As Long, stem As String, parts As Variant
    For Each mark In Array("<", ">", ":", """", "/", "\", "|", "?", "*")
        name = Replace(name, CStr(mark), "_")
    Next mark
    For i = 0 To 31
        name = Replace(name, Chr$(i), "_")
    Next i
    name = Trim$(name)
    Do While Len(name) > 0
        If Right$(name, 1) <> "." And Right$(name, 1) <> " " Then Exit Do
        name = Left$(name, Len(name) - 1)
    Loop
    If Len(name) = 0 Then name = "file"
    parts = Split(name, "."): stem = UCase$(parts(0))
    If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Or stem Like "COM[1-9]" Or stem Like "LPT[1-9]" Then name = "_" & name
    AT_SafeFileName = name
End Function

' @用途: フォルダ選択画面を表示する。キャンセルなら空文字列
' @例: folder = AT_ChooseFolder()
Public Function AT_ChooseFolder(Optional ByVal title As String = "フォルダを選択") As String
    With Application.FileDialog(4)
        .Title = title: .AllowMultiSelect = False
        If .Show = -1 Then AT_ChooseFolder = .SelectedItems(1)
    End With
End Function

' @用途: ファイル選択画面を表示する。キャンセルなら空文字列
' @例: path = AT_ChooseFile("CSV (*.csv),*.csv")
Public Function AT_ChooseFile(Optional ByVal filter As String = "All files (*.*),*.*", Optional ByVal title As String = "ファイルを選択") As String
    Dim result As Variant
    result = Application.GetOpenFilename(FileFilter:=filter, Title:=title, MultiSelect:=False)
    If VarType(result) <> vbBoolean Then AT_ChooseFile = CStr(result)
End Function

' ===== mdlAT_Format =====
' @用途: セルの配置・折り返し・縮小表示を設定する
' @例: AT_SetAlignment ws.Range("A1:D10"), xlLeft, xlCenter, True
Public Sub AT_SetAlignment(ByVal target As Range, Optional ByVal horizontal As Long = xlGeneral, _
    Optional ByVal vertical As Long = xlBottom, Optional ByVal wrap As Boolean = False, Optional ByVal shrink As Boolean = False)
    With target
        .HorizontalAlignment = horizontal: .VerticalAlignment = vertical
        .WrapText = wrap: .ShrinkToFit = shrink
    End With
End Sub

' @用途: 範囲の外枠を設定する
' @例: AT_SetOuterBorder ws.Range("A1:D10"), RGB(0, 0, 0)
Public Sub AT_SetOuterBorder(ByVal target As Range, Optional ByVal color As Long = 0, Optional ByVal weight As Long = xlThin)
    Dim edge As Variant
    For Each edge In Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
        With target.Borders(CLng(edge))
            .LineStyle = xlContinuous: .Weight = weight: .Color = color
        End With
    Next edge
End Sub

' @用途: セルの格子罫線を設定する
' @例: AT_SetGridBorder ws.Range("A1:D10")
Public Sub AT_SetGridBorder(ByVal target As Range, Optional ByVal color As Long = 0, Optional ByVal weight As Long = xlThin)
    With target.Borders
        .LineStyle = xlContinuous: .Weight = weight: .Color = color
    End With
End Sub

' @用途: 塗りつぶしを設定する
' @例: AT_SetFill ws.Range("A1:D1"), RGB(220, 230, 240)
Public Sub AT_SetFill(ByVal target As Range, ByVal color As Long)
    target.Interior.Pattern = xlSolid: target.Interior.Color = color
End Sub

' @用途: フォントをまとめて設定する
' @例: AT_SetFont ws.Range("A1:D10"), "游ゴシック", 11, False
Public Sub AT_SetFont(ByVal target As Range, Optional ByVal name As String = "游ゴシック", _
    Optional ByVal size As Double = 11, Optional ByVal bold As Boolean = False, Optional ByVal color As Long = 0)
    With target.Font
        .Name = name: .Size = size: .Bold = bold: .Color = color
    End With
End Sub

' @用途: 見出し行を太字・背景色付きにする
' @例: AT_FormatHeader ws.Range("A1:F1")
Public Sub AT_FormatHeader(ByVal target As Range, Optional ByVal background As Long = 15790320, Optional ByVal foreground As Long = 0)
    target.Font.Bold = True: target.Font.Color = foreground
    target.Interior.Pattern = xlSolid: target.Interior.Color = background
    target.WrapText = True: target.VerticalAlignment = xlCenter
End Sub

' @用途: 指定範囲のフォント・背景・外枠・表示形式を一括設定する
' @例: AT_FormatBlock ws.Range("B2:D10"), "#,##0", 11
Public Sub AT_FormatBlock(ByVal target As Range, Optional ByVal numberFormat As String = "General", Optional ByVal fontSize As Double = 11)
    AT_SetAlignment target
    AT_SetFont target, "游ゴシック", fontSize
    AT_SetOuterBorder target
    target.NumberFormat = numberFormat
End Sub

' @用途: 指定桁数で四捨五入・切り上げ・切り捨てする
' @例: Debug.Print AT_RoundNumber(12.345, 2, "HALF_UP")
Public Function AT_RoundNumber(ByVal value As Double, Optional ByVal digits As Long = 0, Optional ByVal mode As String = "HALF_UP") As Double
    Select Case UCase$(mode)
        Case "HALF_UP", "ROUND_HALF_UP": AT_RoundNumber = Application.WorksheetFunction.Round(value, digits)
        Case "UP", "ROUND_UP": AT_RoundNumber = Application.WorksheetFunction.RoundUp(value, digits)
        Case "DOWN", "ROUND_DOWN": AT_RoundNumber = Application.WorksheetFunction.RoundDown(value, digits)
        Case Else: Err.Raise 5, "AT_RoundNumber", "Use HALF_UP, UP or DOWN."
    End Select
End Function

' ===== mdlAT_Range =====
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

' ===== mdlAT_Shapes =====
' @用途: 吹き出しを別シートの同じ位置へ作成する
' @例: AT_CopyCallouts sourceSheet, targetSheet
Public Sub AT_CopyCallouts(ByVal source As Worksheet, ByVal destination As Worksheet)
    Dim shape As Shape, copied As Shape
    If source Is destination Then Err.Raise 5, "AT_CopyCallouts", "Source and destination must differ."
    For Each shape In source.Shapes
        If shape.Type = msoAutoShape Then
            Select Case shape.AutoShapeType
                Case msoShapeRoundedRectangularCallout, msoShapeRectangularCallout, msoShapeOvalCallout, msoShapeCloudCallout
                    Set copied = destination.Shapes.AddShape(shape.AutoShapeType, shape.Left, shape.Top, shape.Width, shape.Height)
                    copied.TextFrame2.TextRange.Text = shape.TextFrame2.TextRange.Text
                    With copied.TextFrame2.TextRange.Font
                        .Name = shape.TextFrame2.TextRange.Font.Name
                        .Size = shape.TextFrame2.TextRange.Font.Size
                        .Bold = shape.TextFrame2.TextRange.Font.Bold
                        .Fill.ForeColor.RGB = shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End With
                    copied.Fill.Visible = shape.Fill.Visible
                    copied.Fill.ForeColor.RGB = shape.Fill.ForeColor.RGB
                    copied.Line.Visible = shape.Line.Visible
                    copied.Line.ForeColor.RGB = shape.Line.ForeColor.RGB
                    copied.Line.Weight = shape.Line.Weight
            End Select
        End If
    Next shape
End Sub

' @用途: 指定セルの位置にメモ用テキストボックスを作る
' @例: Set note = AT_AddNoteBox(ws, ws.Range("F2"), "確認事項")
Public Function AT_AddNoteBox(ByVal ws As Worksheet, ByVal anchor As Range, ByVal text As String, _
    Optional ByVal width As Double = 180, Optional ByVal height As Double = 80) As Shape
    Dim result As Shape
    If Not (anchor.Parent Is ws) Then Err.Raise 5, "AT_AddNoteBox", "Anchor belongs to another sheet."
    If width <= 0 Or height <= 0 Then Err.Raise 5, "AT_AddNoteBox", "Invalid dimensions."
    Set result = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, anchor.Left, anchor.Top, width, height)
    result.TextFrame2.TextRange.Text = text
    result.TextFrame2.TextRange.Font.Size = 11
    result.Fill.ForeColor.RGB = RGB(255, 255, 220)
    Set AT_AddNoteBox = result
End Function

' @用途: 特定の接頭辞を持つ図形だけを削除する
' @例: AT_DeleteShapesByPrefix ws, "AT_Note_"
Public Sub AT_DeleteShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    Dim i As Long
    If Len(prefix) = 0 Then Err.Raise 5, "AT_DeleteShapesByPrefix", "Prefix must not be empty."
    For i = ws.Shapes.Count To 1 Step -1
        If StrComp(Left$(ws.Shapes(i).Name, Len(prefix)), prefix, vbTextCompare) = 0 Then ws.Shapes(i).Delete
    Next i
End Sub

' ===== mdlAT_Text =====
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

' ===== mdlAT_Validation =====
' @用途: Null・Empty・空文字か判定する。0・False・セルエラーは空扱いにしない
' @例: Debug.Print AT_IsBlank(ws.Range("A1").Value2)
Public Function AT_IsBlank(ByVal value As Variant) As Boolean
    If IsError(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then AT_IsBlank = True: Exit Function
    If VarType(value) = vbString Then AT_IsBlank = Len(Trim$(CStr(value))) = 0
End Function

' @用途: 必須欄が空のセルをCollectionで返す
' @例: Set missing = AT_CheckRequired(ws.Range("B2:B100"))
Public Function AT_CheckRequired(ByVal scope As Range) As Collection
    Dim result As New Collection, cell As Range
    For Each cell In scope.Cells
        If AT_IsBlank(cell.Value2) Then result.Add cell
    Next cell
    Set AT_CheckRequired = result
End Function

' @用途: 2回目以降に現れる重複値のセルを返す。空欄は既定で無視する
' @例: Set duplicates = AT_CheckDuplicates(ws.Range("A2:A100"))
Public Function AT_CheckDuplicates(ByVal scope As Range, Optional ByVal ignoreBlank As Boolean = True) As Collection
    Dim result As New Collection, seen As Object, cell As Range, value As Variant, key As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For Each cell In scope.Cells
        value = cell.Value2
        If IsError(value) Then
            result.Add cell
        ElseIf Not (ignoreBlank And AT_IsBlank(value)) Then
            If IsNull(value) Then key = "" Else key = CStr(value)
            If seen.Exists(key) Then result.Add cell Else seen.Add key, True
        End If
    Next cell
    Set AT_CheckDuplicates = result
End Function

' @用途: 数値範囲に入っていないセルを返す。上下限を含む
' @例: Set invalid = AT_CheckNumberRange(ws.Range("C2:C100"), 0, 100)
Public Function AT_CheckNumberRange(ByVal scope As Range, ByVal minimum As Double, ByVal maximum As Double, Optional ByVal ignoreBlank As Boolean = True) As Collection
    Dim result As New Collection, cell As Range, value As Variant, valid As Boolean, number As Double
    If minimum > maximum Then Err.Raise 5, "AT_CheckNumberRange", "Minimum exceeds maximum."
    For Each cell In scope.Cells
        value = cell.Value2
        If Not (ignoreBlank And AT_IsBlank(value)) Then
            valid = False
            If Not IsError(value) And Not IsNull(value) And Not IsEmpty(value) Then
                If VarType(value) <> vbBoolean And IsNumeric(value) Then
                    On Error Resume Next
                    Err.Clear: number = CDbl(value)
                    If Err.Number = 0 Then valid = number >= minimum And number <= maximum
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            If Not valid Then result.Add cell
        End If
    Next cell
    Set AT_CheckNumberRange = result
End Function

' @用途: 入力規則のリストを設定する。名前付き範囲か短い候補一覧を渡す
' @例: AT_SetListValidation ws.Range("B2:B100"), "=ChoiceList"
Public Sub AT_SetListValidation(ByVal target As Range, ByVal formula As String, Optional ByVal allowBlank As Boolean = True)
    If Len(formula) = 0 Or Len(formula) > 255 Then Err.Raise 5, "AT_SetListValidation", "Use a named range for long lists."
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formula
        .IgnoreBlank = allowBlank: .InCellDropdown = True: .ShowError = True
    End With
End Sub

' @用途: 値のあるデータセルへ整数の入力規則を設定する
' @例: AT_SetIntegerValidation ws.Range("C2:C100"), 0, 999
Public Sub AT_SetIntegerValidation(ByVal target As Range, ByVal minimum As Long, ByVal maximum As Long, Optional ByVal allowBlank As Boolean = True)
    If minimum > maximum Then Err.Raise 5, "AT_SetIntegerValidation", "Minimum exceeds maximum."
    With target.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=minimum, Formula2:=maximum
        .IgnoreBlank = allowBlank: .ShowError = True
    End With
End Sub

' @用途: セルの編集可否を設定する。シート保護が有効なときに適用される
' @例: AT_SetLocked ws.Range("B2:D100"), False
Public Sub AT_SetLocked(ByVal target As Range, ByVal locked As Boolean)
    target.Locked = locked
End Sub
