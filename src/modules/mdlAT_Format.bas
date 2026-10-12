Attribute VB_Name = "mdlAT_Format"
Option Explicit


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
