Attribute VB_Name = "mdlAT_SelfTest"
Option Explicit

' Import all modules and classes, then run this in Excel. Uses a disposable workbook.
Public Sub AT_RunSelfTests()
    Dim book As Workbook, ws As Worksheet, logs As Worksheet, guard As New clsAT_AppGuard
    Dim rule As New clsAT_Rule, logger As New clsAT_Logger
    Dim data As Variant, parsed As Variant, index As Object, result As Collection, cell As Range
    Dim number As Double, stamp As Date, i As Long, n As Long, message As String, state As AT_ExcelState
    Dim originalEvents As Boolean, originalAlerts As Boolean, originalUpdating As Boolean, originalCalculation As Long
    originalEvents = Application.EnableEvents: originalAlerts = Application.DisplayAlerts
    originalUpdating = Application.ScreenUpdating: originalCalculation = Application.Calculation
    On Error GoTo Failed
    guard.Start
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = book.Worksheets(1)
    AT_TestAssert AT_IsBlank(Empty), "Empty"
    AT_TestAssert AT_IsBlank(Null), "Null"
    AT_TestAssert Not AT_IsBlank(0), "Zero is not blank"
    AT_TestAssert Not AT_IsBlank(False), "False is not blank"
    AT_TestAssert Not AT_IsBlank(CVErr(xlErrNA)), "Error is not blank"
    AT_TestAssert AT_TextOr(Null, "missing") = "missing", "Text fallback"
    AT_TestAssert AT_ReplaceText("abc", "", "x") = "abc", "Empty replacement pattern"
    AT_TestAssert AT_NormalizeSpaces("  A" & vbTab & "B  ") = "A B", "Whitespace"
    AT_TestAssert AT_Between("x[ABC]y", "[", "]") = "ABC", "Between"
    AT_TestAssert AT_NameFromLabel("1||A||B", "||") = "A||B", "Variable delimiter length"
    AT_TestAssert AT_CodeFromLabel("ABC", "||") = "ABC", "Missing delimiter"
    AT_TestAssert AT_BrToNewline("A<BR />B") = "A" & vbLf & "B", "BR variants"
    AT_TestAssert AT_PadLeft("7", 3) = "007", "Padding"
    AT_TestAssert AT_TryNumber("12.5", number) And number = 12.5, "Number"
    AT_TestAssert Not AT_TryNumber("bad", number), "Invalid number"
    AT_TestAssert Not AT_TryNumber(Empty, number), "Missing number"
    For i = 1 To 16384
        AT_TestAssert AT_ColumnNumber(AT_ColumnLetter(i)) = i, "Column round trip: " & i
    Next i
    AT_TestAssert AT_ArrayLength(Array()) = 0, "Empty array"
    data = AT_UniqueValues(Array("A", "a", "B"))
    AT_TestAssert AT_ArrayLength(data) = 2, "Case-insensitive uniqueness"
    data = AT_FilterNonEmpty(Array("", 0, "A", Null))
    AT_TestAssert AT_ArrayLength(data) = 2, "Filter preserves zero"
    AT_TestAssert AT_SafeSheetName("A/B:C") = "A_B_C", "Sheet name"
    AT_TestAssert AT_SafeSheetName("''") = "Sheet", "Empty sheet name"
    AT_TestAssert AT_SafeFileName("CON.txt") = "_CON.txt", "Reserved filename"
    AT_TestAssert AT_TryIsoDate("2024-02-29", stamp), "Leap date"
    AT_TestAssert Not AT_TryIsoDate("2025-02-29", stamp), "Invalid leap date"
    AT_TestAssert AT_FormatCompactDate("202610121530") = "2026/10/12 15:30", "Compact date"
    AT_TestAssert AT_WeekMonday(DateSerial(2026, 10, 14)) = DateSerial(2026, 10, 12), "Week start"
    AT_TestAssert AT_RoundNumber(12, 0, "UP") = 12, "No artificial increment"
    AT_TestAssert AT_RoundNumber(-12.1, 0, "UP") = -13, "Round away from zero"
    AT_TestAssert AT_LastRow(ws) = 0, "Empty last row"
    ws.Range("A1").Value2 = "A*": ws.Range("A2").Value2 = "ABC"
    ws.Range("A20").Formula = "="""""
    AT_TestAssert AT_LastRow(ws, "A") = 20, "Formula counts in last row"
    Set cell = AT_FindCell(ws.Range("A1:A2"), "A*")
    AT_TestAssert Not cell Is Nothing, "Literal wildcard found"
    AT_TestAssert cell.Row = 1, "Literal wildcard location"
    AT_TestAssert AT_HeaderColumn(ws, "A*") = 1, "Header lookup"
    data = AT_ReadMatrix(ws.Range("A2"))
    AT_TestAssert data(1, 1) = "ABC", "Single-cell matrix"
    Set cell = AT_UnionRange(Nothing, ws.Range("A1"))
    AT_TestAssert cell.Address = "$A$1", "Union with Nothing"
    ws.Range("B2").Value2 = 0: ws.Range("B3").Value2 = "": ws.Range("B4").Value2 = "A"
    Set result = AT_CheckRequired(ws.Range("B2:B4"))
    AT_TestAssert result.Count = 1, "Required zero handling"
    rule.Configure ">", 10, True
    AT_TestAssert Not rule.Check("2"), "Numeric comparison"
    AT_TestAssert rule.Check("11"), "Numeric comparison success"
    AT_TestAssert Not rule.Check("bad"), "Invalid numeric value"
    rule.Configure "IS"
    AT_TestAssert rule.Check(Null), "Rule null handling"
    parsed = AT_ParseCsv("A,B" & vbCrLf & """x,y"",""a""""b""" & vbCrLf)
    AT_TestAssert parsed(2, 1) = "x,y", "CSV comma"
    AT_TestAssert parsed(2, 2) = "a""b", "CSV escaped quote"
    data = AT_ParseCsv(AT_MatrixToCsv(parsed))
    AT_TestAssert data(2, 2) = parsed(2, 2), "CSV round trip"
    parsed = AT_ParseCsv("""A" & vbLf & "B"",0")
    AT_TestAssert parsed(1, 1) = "A" & vbLf & "B", "CSV quoted newline"
    parsed = AT_ParseCsv("A,")
    AT_TestAssert UBound(parsed, 2) = 2 And parsed(1, 2) = "", "CSV trailing empty field"
    parsed = AT_ParseCsv("A,1" & vbLf & "B,2" & vbLf & "A,3")
    Set index = AT_GroupSum(parsed, 1, 2)
    AT_TestAssert index("A") = 4, "Grouped totals"
    Set index = AT_IndexRows(parsed, 1, "LAST")
    AT_TestAssert index("A") = 3, "Duplicate last policy"
    Set logs = book.Worksheets.Add
    logger.Attach logs: logger.WriteEntry "INFO", "=1+1", "test"
    AT_TestAssert Not logs.Range("D2").HasFormula, "Log is literal text"
    state = AT_PauseExcel()
    AT_RestoreExcel state
    AT_TestAssert Not state.Captured, "State capture released"
    book.Close SaveChanges:=False
    Set book = Nothing
    guard.Restore
    AT_TestAssert Application.EnableEvents = originalEvents, "Events restored"
    AT_TestAssert Application.DisplayAlerts = originalAlerts, "Alerts restored"
    AT_TestAssert Application.ScreenUpdating = originalUpdating, "Updating restored"
    AT_TestAssert Application.Calculation = originalCalculation, "Calculation restored"
    Debug.Print "PASS: AT_RunSelfTests"
    Exit Sub
Failed:
    n = Err.Number: message = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    guard.Restore
    On Error GoTo 0
    Err.Raise n, "AT_RunSelfTests", message
End Sub

Private Sub AT_TestAssert(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 700, "AT_RunSelfTests", label
End Sub
