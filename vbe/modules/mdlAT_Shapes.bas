Attribute VB_Name = "mdlAT_Shapes"
Option Explicit


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
