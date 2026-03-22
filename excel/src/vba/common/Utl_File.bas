Option Explicit

' ============================================
' Module   : Utl_File
' Layer    : Common / Utility
' Purpose  : File operations using FileSystemObject
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Ported from FlowBase Utl_File
' ============================================

' ============================================
' CreateFolder
' Create folder if it doesn't exist (recursive)
' ============================================
Public Function CreateFolder(folderPath As String) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        CreateFolder = True
        Exit Function
    End If

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        If Not CreateFolder(parentPath) Then
            CreateFolder = False
            Exit Function
        End If
    End If

    fso.CreateFolder folderPath
    CreateFolder = True
    Exit Function

ErrHandler:
    CreateFolder = False
End Function

' ============================================
' WriteTextFile
' Write text content to file (UTF-8 with BOM)
' ============================================
Public Function WriteTextFile(filePath As String, content As String, Optional append As Boolean = False) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(filePath)

    If Not CreateFolder(parentPath) Then
        WriteTextFile = False
        Exit Function
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 2  ' adTypeText
        .Charset = "UTF-8"
        .Open

        If append And fso.FileExists(filePath) Then
            Dim existingStream As Object
            Set existingStream = CreateObject("ADODB.Stream")
            existingStream.Type = 2
            existingStream.Charset = "UTF-8"
            existingStream.Open
            existingStream.LoadFromFile filePath
            Dim existing As String
            existing = existingStream.ReadText
            existingStream.Close
            Set existingStream = Nothing
            .WriteText existing
        End If

        .WriteText content
        .SaveToFile filePath, 2  ' adSaveCreateOverWrite
        .Close
    End With

    Set stream = Nothing
    WriteTextFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    WriteTextFile = False
End Function

' ============================================
' ReadTextFile
' Read text content from file (UTF-8)
' ============================================
Public Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        ReadTextFile = ""
        Exit Function
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadTextFile = .ReadText
        .Close
    End With

    Set stream = Nothing
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    ReadTextFile = ""
End Function

' ============================================
' FileExists
' ============================================
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

' ============================================
' SanitizeFilename
' Remove invalid characters from filename
' ============================================
Public Function SanitizeFilename(filename As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim invalidChars As String

    invalidChars = "<>:""/\|?*"
    result = filename

    For i = 1 To Len(invalidChars)
        char = Mid(invalidChars, i, 1)
        result = Replace(result, char, "_")
    Next i

    SanitizeFilename = result
End Function

' ============================================
' BuildFilePath
' Combine folder path and filename
' ============================================
Public Function BuildFilePath(folderPath As String, filename As String) As String
    If Right(folderPath, 1) = "\" Then
        BuildFilePath = folderPath & filename
    Else
        BuildFilePath = folderPath & "\" & filename
    End If
End Function
