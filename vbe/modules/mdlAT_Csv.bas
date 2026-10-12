Attribute VB_Name = "mdlAT_Csv"
Option Explicit


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
