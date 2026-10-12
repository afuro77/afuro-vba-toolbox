Attribute VB_Name = "mdlAT_Recipes"
Option Explicit

' 必要部品: clsAT_AppGuard / mdlAT_Range / mdlAT_Format / mdlAT_Dictionary
' 呼び出し例: AT_RecipeSummarize source.Range("A2:C100"), target.Range("A1")
Public Sub AT_RecipeSummarize(ByVal source As Range, ByVal destination As Range)
    Dim guard As New clsAT_AppGuard, data As Variant, totals As Object, output As Variant
    Dim number As Long, message As String
    On Error GoTo Failed
    guard.Start
    data = AT_ReadMatrix(source)
    Set totals = AT_GroupSum(data, 1, 3)
    output = AT_DictionaryMatrix(totals)
    If Not IsEmpty(output) Then
        destination.Resize(1, 2).Value = Array("キー", "合計")
        AT_FormatHeader destination.Resize(1, 2)
        AT_WriteMatrix destination.Offset(1, 0), output
    End If
    guard.Restore
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    guard.Restore
    On Error GoTo 0
    Err.Raise number, "AT_RecipeSummarize", message
End Sub

' 必要部品: mdlAT_Files / mdlAT_Range / mdlAT_Text / mdlAT_Csv
' 呼び出し例: AT_RecipeImportCsv inputPath, target.Range("A1")
' CSVのフィールドを文字列として書き込み、先頭ゼロを保つ。
Public Sub AT_RecipeImportCsv(ByVal path As String, ByVal destination As Range)
    Dim data As Variant
    data = AT_ParseCsv(AT_ReadTextFile(path, "utf-8"))
    If IsEmpty(data) Then Exit Sub
    destination.Resize(UBound(data, 1), UBound(data, 2)).NumberFormat = "@"
    AT_WriteMatrix destination, data
End Sub

' 必要部品: mdlAT_Validation / mdlAT_Range / mdlAT_Format
' 呼び出し例: AT_RecipeHighlightMissing ws.Range("B2:B100")
Public Sub AT_RecipeHighlightMissing(ByVal scope As Range)
    Dim missing As Collection, cell As Range, marked As Range
    Set missing = AT_CheckRequired(scope)
    For Each cell In missing
        Set marked = AT_UnionRange(marked, cell)
    Next cell
    If Not marked Is Nothing Then AT_SetFill marked, RGB(255, 230, 230)
End Sub
