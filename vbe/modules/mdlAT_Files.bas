Attribute VB_Name = "mdlAT_Files"
Option Explicit


' @用途: フルパスを結合する。後半は相対パスを指定する
' @例: path = AT_PathJoin("C:\sample", "data\input.csv")
Public Function AT_PathJoin(ByVal folder As String, ByVal child As String) As String
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Left$(child, 1) = "\" Or InStr(child, ":") > 0 Then Err.Raise 5, "AT_PathJoin", "Child must be a relative path."
    AT_PathJoin = fs.BuildPath(folder, child)
End Function

' @用途: ファイルの存在を確認する
' @例: If AT_FileExists(path) Then Debug.Print path
Public Function AT_FileExists(ByVal path As String) As Boolean
    AT_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' @用途: フォルダの存在を確認する
' @例: Debug.Print AT_FolderExists("C:\sample")
Public Function AT_FolderExists(ByVal path As String) As Boolean
    AT_FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(path)
End Function

' @用途: 親フォルダも含めてフォルダを作る。既存なら何もしない
' @例: AT_EnsureFolder "C:\sample\output\2026"
Public Sub AT_EnsureFolder(ByVal path As String)
    Dim fs As Object, parent As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    path = fs.GetAbsolutePathName(path)
    If fs.FolderExists(path) Then Exit Sub
    parent = fs.GetParentFolderName(path)
    If Len(parent) = 0 Or parent = path Then Err.Raise 76, "AT_EnsureFolder", "Root folder is unavailable."
    AT_EnsureFolder parent
    fs.CreateFolder path
End Sub

' @用途: ファイル名をフルパスから取得する
' @例: Debug.Print AT_FileName("C:\sample\input.csv")
Public Function AT_FileName(ByVal path As String) As String
    AT_FileName = CreateObject("Scripting.FileSystemObject").GetFileName(path)
End Function

' @用途: 拡張子なしのファイル名を取得する
' @例: Debug.Print AT_BaseName("C:\sample\input.csv")
Public Function AT_BaseName(ByVal path As String) As String
    AT_BaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(path)
End Function

' @用途: 拡張子を小文字で取得する。先頭のドットは含めない
' @例: Debug.Print AT_Extension("input.CSV")
Public Function AT_Extension(ByVal path As String) As String
    AT_Extension = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(path))
End Function

' @用途: ファイルの親フォルダを取得する
' @例: Debug.Print AT_ParentFolder("C:\sample\input.csv")
Public Function AT_ParentFolder(ByVal path As String) As String
    AT_ParentFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(path)
End Function

' @用途: 条件に一致するファイルをCollectionで取得する。再帰は任意
' @例: Set files = AT_ListFiles("C:\sample", "*.csv", True)
Public Function AT_ListFiles(ByVal folder As String, Optional ByVal pattern As String = "*", Optional ByVal recursive As Boolean = False) As Collection
    Dim result As New Collection, fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    AT_FilesCollect fs.GetFolder(folder), pattern, recursive, result
    Set AT_ListFiles = result
End Function

Private Sub AT_FilesCollect(ByVal folder As Object, ByVal pattern As String, ByVal recursive As Boolean, ByVal result As Collection)
    Dim file As Object, child As Object
    For Each file In folder.Files
        If LCase$(file.Name) Like LCase$(pattern) Then
            If Left$(file.Name, 2) <> "~$" Then result.Add file.Path
        End If
    Next file
    If recursive Then
        For Each child In folder.SubFolders
            If (child.Attributes And 1024) = 0 Then AT_FilesCollect child, pattern, True, result
        Next child
    End If
End Sub

' @用途: 既存ファイルと重ならないパスを作る。ファイル自体は作成しない
' @例: output = AT_UniqueFilePath("C:\sample\output.csv")
Public Function AT_UniqueFilePath(ByVal path As String) As String
    Dim fs As Object, stem As String, extension As String, candidate As String, i As Long
    Set fs = CreateObject("Scripting.FileSystemObject")
    stem = fs.BuildPath(fs.GetParentFolderName(path), fs.GetBaseName(path))
    extension = fs.GetExtensionName(path)
    If Len(extension) > 0 Then extension = "." & extension
    candidate = path
    Do While fs.FileExists(candidate) Or fs.FolderExists(candidate)
        i = i + 1: candidate = stem & "_" & i & extension
    Loop
    AT_UniqueFilePath = candidate
End Function

' @用途: ファイルをコピーする。既定では上書きしない
' @例: AT_CopyFile sourcePath, destinationPath
Public Sub AT_CopyFile(ByVal source As String, ByVal destination As String, Optional ByVal overwrite As Boolean = False)
    CreateObject("Scripting.FileSystemObject").CopyFile source, destination, overwrite
End Sub

' @用途: テキストファイルを指定文字コードで読む。既定はUTF-8
' @例: text = AT_ReadTextFile(path, "utf-8")
Public Function AT_ReadTextFile(ByVal path As String, Optional ByVal charset As String = "utf-8") As String
    Dim stream As Object, number As Long, message As String
    Set stream = CreateObject("ADODB.Stream")
    On Error GoTo Failed
    stream.Type = 2: stream.Charset = charset: stream.Open
    stream.LoadFromFile path
    AT_ReadTextFile = stream.ReadText(-1)
    stream.Close
    Exit Function
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    stream.Close
    On Error GoTo 0
    Err.Raise number, "AT_ReadTextFile", message
End Function

' @用途: UTF-8でテキストを書き出す。既定はBOMなし・上書き禁止
' @例: AT_WriteUtf8 path, text
Public Sub AT_WriteUtf8(ByVal path As String, ByVal text As String, Optional ByVal overwrite As Boolean = False, Optional ByVal includeBom As Boolean = False)
    Dim stream As Object, binary As Object, number As Long, message As String, saveMode As Long
    Set stream = CreateObject("ADODB.Stream")
    On Error GoTo Failed
    If overwrite Then saveMode = 2 Else saveMode = 1
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText text
    If includeBom Then
        stream.SaveToFile path, saveMode
    Else
        stream.Position = 0: stream.Type = 1
        Set binary = CreateObject("ADODB.Stream")
        binary.Type = 1: binary.Open
        If stream.Size >= 3 Then
            stream.Position = 3
            If stream.Size > 3 Then binary.Write stream.Read
        End If
        binary.SaveToFile path, saveMode
        binary.Close
    End If
    stream.Close
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    stream.Close
    If Not binary Is Nothing Then binary.Close
    On Error GoTo 0
    Err.Raise number, "AT_WriteUtf8", message
End Sub

' @用途: Windowsのファイル名に使えない文字を置換する
' @例: name = AT_SafeFileName("売上:2026/10.csv")
Public Function AT_SafeFileName(ByVal name As String) As String
    Dim mark As Variant, i As Long, stem As String, parts As Variant
    For Each mark In Array("<", ">", ":", """", "/", "\", "|", "?", "*")
        name = Replace(name, CStr(mark), "_")
    Next mark
    For i = 0 To 31
        name = Replace(name, Chr$(i), "_")
    Next i
    name = Trim$(name)
    Do While Len(name) > 0
        If Right$(name, 1) <> "." And Right$(name, 1) <> " " Then Exit Do
        name = Left$(name, Len(name) - 1)
    Loop
    If Len(name) = 0 Then name = "file"
    parts = Split(name, "."): stem = UCase$(parts(0))
    If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Or stem Like "COM[1-9]" Or stem Like "LPT[1-9]" Then name = "_" & name
    AT_SafeFileName = name
End Function

' @用途: フォルダ選択画面を表示する。キャンセルなら空文字列
' @例: folder = AT_ChooseFolder()
Public Function AT_ChooseFolder(Optional ByVal title As String = "フォルダを選択") As String
    With Application.FileDialog(4)
        .Title = title: .AllowMultiSelect = False
        If .Show = -1 Then AT_ChooseFolder = .SelectedItems(1)
    End With
End Function

' @用途: ファイル選択画面を表示する。キャンセルなら空文字列
' @例: path = AT_ChooseFile("CSV (*.csv),*.csv")
Public Function AT_ChooseFile(Optional ByVal filter As String = "All files (*.*),*.*", Optional ByVal title As String = "ファイルを選択") As String
    Dim result As Variant
    result = Application.GetOpenFilename(FileFilter:=filter, Title:=title, MultiSelect:=False)
    If VarType(result) <> vbBoolean Then AT_ChooseFile = CStr(result)
End Function
