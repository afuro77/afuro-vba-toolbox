Attribute VB_Name = "mdlAT_Excel"
Option Explicit


Public Type AT_ExcelState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As Long
    StatusBar As Variant
    Captured As Boolean
End Type

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
