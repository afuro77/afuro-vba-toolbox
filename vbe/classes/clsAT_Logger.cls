VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsAT_Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Worksheet logger. Bind the target explicitly; no ActiveSheet or ThisWorkbook.
Private destination As Worksheet

' @用途: ログ用シートを設定する。空なら見出しを作る
' @例: Dim logger As New clsAT_Logger: logger.Attach logSheet
Public Sub Attach(ByVal ws As Worksheet)
    Dim headers(1 To 1, 1 To 4) As Variant
    If Application.CountA(ws.UsedRange) = 0 Then
        headers(1, 1) = "日時": headers(1, 2) = "レベル": headers(1, 3) = "処理名": headers(1, 4) = "内容"
        ws.Range("A1:D1").Value2 = headers
        ws.Range("A1:D1").Font.Bold = True
    ElseIf ws.Range("A1").Value2 <> "日時" Or ws.Range("D1").Value2 <> "内容" Then
        Err.Raise 5, "clsAT_Logger.Attach", "Choose a blank sheet or an existing log sheet."
    End If
    Set destination = ws
End Sub

' @用途: 日時付きのログを追記する。文字列は数式として実行しない
' @例: logger.WriteEntry "INFO", "処理完了", "集計"
Public Sub WriteEntry(ByVal level As String, ByVal message As String, Optional ByVal procedure As String = "")
    Dim last As Range, row As Long, values(1 To 1, 1 To 3) As Variant
    If destination Is Nothing Then Err.Raise 5, "clsAT_Logger.Write", "Attach a log sheet first."
    Set last = destination.Range("A:D").Find(What:="*", After:=destination.Range("A1"), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    row = 2
    If Not last Is Nothing Then row = last.Row + 1
    If row > destination.Rows.Count Then Err.Raise 6, "clsAT_Logger.Write", "Log sheet is full."
    destination.Cells(row, 1).Value = Now
    destination.Cells(row, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    values(1, 1) = level: values(1, 2) = procedure: values(1, 3) = message
    destination.Cells(row, 2).Resize(1, 3).NumberFormat = "@"
    destination.Cells(row, 2).Resize(1, 3).Value2 = values
End Sub

' @用途: エラー番号と内容をログに残す
' @例: logger.WriteError Err.Number, Err.Description, "集計"
Public Sub WriteError(ByVal number As Long, ByVal description As String, Optional ByVal procedure As String = "")
    WriteEntry "ERROR", "[" & CStr(number) & "] " & description, procedure
End Sub
