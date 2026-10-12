VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsAT_AppGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Excel settings guard. Standalone class; no standard module required.
Private updating As Boolean, events As Boolean, alerts As Boolean
Private calculation As Long, status As Variant, captured As Boolean

' @用途: 高速化を開始する。クラス単体で使え、元の設定を保持する
' @例: Dim guard As New clsAT_AppGuard: guard.Start
Public Sub Start()
    Dim number As Long, message As String
    If captured Then Err.Raise 5, "clsAT_AppGuard", "Already active."
    updating = Application.ScreenUpdating: events = Application.EnableEvents
    alerts = Application.DisplayAlerts: calculation = Application.Calculation
    status = Application.StatusBar: captured = True
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    Restore
    On Error GoTo 0
    Err.Raise number, "clsAT_AppGuard.Start", message
End Sub

' @用途: 設定を明示的に復元する。参照解放時にも復元を試みる
' @例: guard.Restore: Set guard = Nothing
Public Sub Restore()
    Dim number As Long, message As String
    If Not captured Then Exit Sub
    On Error Resume Next
    Application.Calculation = calculation
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.DisplayAlerts = alerts
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.EnableEvents = events
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.ScreenUpdating = updating
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    Application.StatusBar = status
    If Err.Number <> 0 Then number = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "clsAT_AppGuard.Restore", message
    captured = False
End Sub

Public Property Get Active() As Boolean
    Active = captured
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    Restore
End Sub
