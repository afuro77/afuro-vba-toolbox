Attribute VB_Name = "mdlAT_Validation"
Option Explicit


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
