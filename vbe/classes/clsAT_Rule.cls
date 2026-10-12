VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsAT_Rule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Reusable comparison rule. No workbook-specific configuration or dependencies.
Private operation As String, expected As Variant, numeric As Boolean, compare As Long

' @ópìr: î‰ärèåèÇê›íËÇ∑ÇÈÅBêîílî‰ärÇ©ï∂éöóÒî‰ärÇ©Çñæé¶Ç∑ÇÈ
' @ó·: Dim rule As New clsAT_Rule: rule.Configure ">=", 0, True
Public Sub Configure(ByVal comparison As String, Optional ByVal expectedValue As Variant = "", _
    Optional ByVal numericComparison As Boolean = False, Optional ByVal matchCase As Boolean = False)
    comparison = UCase$(Trim$(comparison))
    Select Case comparison
        Case "=", "<>", "!=", ">", ">=", "<", "<=", "IS", "IS NOT"
        Case Else: Err.Raise 5, "clsAT_Rule.Configure", "Unsupported comparison."
    End Select
    If comparison <> "IS" And comparison <> "IS NOT" Then
        If IsError(expectedValue) Or IsNull(expectedValue) Or IsObject(expectedValue) Or IsArray(expectedValue) Then Err.Raise 5, "clsAT_Rule.Configure", "Invalid expected value."
        If numericComparison Then
            If IsEmpty(expectedValue) Or VarType(expectedValue) = vbBoolean Or Not IsNumeric(expectedValue) Then Err.Raise 5, "clsAT_Rule.Configure", "Expected value must be numeric."
            expectedValue = CDbl(expectedValue)
        End If
    End If
    operation = comparison: expected = expectedValue: numeric = numericComparison
    If matchCase Then compare = vbBinaryCompare Else compare = vbTextCompare
End Sub

' @ópìr: 1Ç¬ÇÃílÇ™èåèÇñûÇΩÇ∑Ç©îªíËÇ∑ÇÈÅBÉGÉâÅ[ílÇ‚å^ïsàÍívÇÕFalse
' @ó·: Debug.Print rule.Check(ws.Range("C2").Value2)
Public Function Check(ByVal value As Variant) As Boolean
    Dim relation As Long, blank As Boolean, number As Double
    If Len(operation) = 0 Then Err.Raise 5, "clsAT_Rule.Check", "Configure the rule first."
    If IsError(value) Or IsObject(value) Or IsArray(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then
        blank = True
    ElseIf VarType(value) = vbString Then
        blank = Len(Trim$(CStr(value))) = 0
    End If
    If operation = "IS" Then Check = blank: Exit Function
    If operation = "IS NOT" Then Check = Not blank: Exit Function
    If IsNull(value) Then Exit Function
    If numeric Then
        If blank Or VarType(value) = vbBoolean Then Exit Function
        If Not IsNumeric(value) Then Exit Function
        On Error GoTo InvalidValue
        number = CDbl(value)
        If number < CDbl(expected) Then
            relation = -1
        ElseIf number > CDbl(expected) Then
            relation = 1
        Else
            relation = 0
        End If
    Else
        relation = StrComp(CStr(value), CStr(expected), compare)
    End If
    Select Case operation
        Case "=": Check = relation = 0
        Case "<>", "!=": Check = relation <> 0
        Case ">": Check = relation > 0
        Case ">=": Check = relation >= 0
        Case "<": Check = relation < 0
        Case "<=": Check = relation <= 0
    End Select
InvalidValue:
End Function

' @ópìr: èåèÇñûÇΩÇ≥Ç»Ç¢ÉZÉãÇCollectionÇ≈éÊìæÇ∑ÇÈ
' @ó·: Set invalid = rule.FindInvalid(ws.Range("C2:C100"))
Public Function FindInvalid(ByVal scope As Range) As Collection
    Dim result As New Collection, cell As Range
    For Each cell In scope.Cells
        If Not Check(cell.Value2) Then result.Add cell
    Next cell
    Set FindInvalid = result
End Function
