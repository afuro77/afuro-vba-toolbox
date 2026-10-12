Attribute VB_Name = "mdlAT_Date"
Option Explicit


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
