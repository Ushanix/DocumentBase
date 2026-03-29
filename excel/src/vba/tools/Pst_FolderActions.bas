Attribute VB_Name = "Pst_FolderActions"
Option Explicit

' ============================================
' Module   : Pst_FolderActions
' Layer    : Presentation
' Purpose  : Folder selection and Explorer integration for output path
' Version  : 1.0.0
' Created  : 2026-03-29
' ============================================

Private Const TOOL_NAME As String = "FolderActions"

' ============================================
' SelectOutputFolder
' Open folder picker dialog and save selected path to
' header_info.folder_output_path on current DOC- sheet.
' ============================================
Public Sub SelectOutputFolder()
    On Error GoTo EH

    ' Check current sheet is DOC-*
    Dim currentSheet As String
    currentSheet = ActiveSheet.Name

    If Left(currentSheet, Len(PREFIX_COLLECTION)) <> PREFIX_COLLECTION Then
        MsgBox "Please run from a DOC-* sheet.", vbExclamation, "Error"
        Exit Sub
    End If

    If Left(currentSheet, Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
        MsgBox "Cannot set output path on template sheet.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(currentSheet)

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)

    ' Resolve initial folder: folder_output_path > ThisWorkbook.Path
    Dim initialDir As String
    initialDir = ThisWorkbook.Path

    Dim headerInfo As Object
    Set headerInfo = Nothing

    If markerRow > 0 Then
        Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

        If headerInfo.Exists("folder_output_path") Then
            Dim currentPath As String
            currentPath = Trim(CStr(headerInfo("folder_output_path")))
            If Len(currentPath) > 0 Then
                initialDir = currentPath
            End If
        End If
    End If

    ' Build collection folder name
    Dim collectionId As String
    Dim collectionName As String
    collectionId = ""
    collectionName = ""

    If Not headerInfo Is Nothing Then
        If headerInfo.Exists("collection_id") Then
            collectionId = CStr(headerInfo("collection_id"))
        End If
        If headerInfo.Exists("collection_name") Then
            collectionName = CStr(headerInfo("collection_name"))
        End If
    End If

    Dim collFolderName As String
    If Len(collectionName) > 0 Then
        collFolderName = SanitizeFilename(collectionId & "_" & collectionName)
    Else
        collFolderName = SanitizeFilename(collectionId)
    End If

    ' Show folder picker (user selects parent directory)
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .title = "Select parent folder for collection output"
        .InitialFileName = initialDir

        If .Show <> -1 Then
            Exit Sub
        End If

        Dim selectedPath As String
        selectedPath = .SelectedItems(1)
    End With

    ' Ensure path ends with collection folder
    Dim finalPath As String

    ' Check if selected path already ends with collection folder name
    Dim selectedFolderName As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    selectedFolderName = fso.GetFileName(selectedPath)

    If UCase(selectedFolderName) = UCase(collFolderName) Then
        ' Already pointing to collection folder
        finalPath = selectedPath
    Else
        ' Append collection folder
        finalPath = BuildFilePath(selectedPath, collFolderName)
    End If

    ' Create folder if needed
    If Not FolderExists(finalPath) Then
        If Not CreateFolder(finalPath) Then
            MsgBox "Failed to create folder:" & vbCrLf & finalPath, vbExclamation, "Error"
            Exit Sub
        End If
        LogInfo TOOL_NAME, "Created collection folder: " & finalPath
    End If

    ' Update header_info
    If markerRow > 0 Then
        If UpdateKeyValueTable(ws, markerRow + 1, "folder_output_path", finalPath) Then
            LogInfo TOOL_NAME, "folder_output_path set: " & finalPath
            MsgBox "Output path set:" & vbCrLf & finalPath, vbInformation, "Complete"
        Else
            MsgBox "folder_output_path key not found in HeaderInfo." & vbCrLf & _
                   "Please add the key first.", vbExclamation, "Error"
        End If
    End If

    Exit Sub

EH:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' OpenOutputFolder
' Open the resolved output folder in Windows Explorer.
' Calls Pst_OutputNotes.ResolveOutputDir to get the same path as OutputNotes.
' ============================================
Public Sub OpenOutputFolder()
    On Error GoTo EH

    ' Check current sheet is DOC-*
    Dim currentSheet As String
    currentSheet = ActiveSheet.Name

    If Left(currentSheet, Len(PREFIX_COLLECTION)) <> PREFIX_COLLECTION Then
        MsgBox "Please run from a DOC-* sheet.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(currentSheet)

    ' Read header_info
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)

    If markerRow = 0 Then
        MsgBox "HeaderInfo not found.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim headerInfo As Object
    Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

    Dim collectionId As String
    If headerInfo.Exists("collection_id") Then
        collectionId = CStr(headerInfo("collection_id"))
    Else
        collectionId = currentSheet
    End If

    ' Use same path resolution as OutputNotes
    Dim outputDir As String
    outputDir = ResolveOutputDir(headerInfo, collectionId)

    If Len(outputDir) = 0 Then
        MsgBox "No output path configured.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Check folder exists
    If Not FolderExists(outputDir) Then
        Dim createResult As VbMsgBoxResult
        createResult = MsgBox("Folder does not exist:" & vbCrLf & outputDir & vbCrLf & vbCrLf & _
                              "Create it?", vbYesNo + vbQuestion, "Folder not found")
        If createResult = vbYes Then
            If Not CreateFolder(outputDir) Then
                MsgBox "Failed to create folder.", vbExclamation, "Error"
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' Open in Explorer
    Shell "explorer.exe """ & outputDir & """", vbNormalFocus
    LogInfo TOOL_NAME, "Opened folder: " & outputDir

    Exit Sub

EH:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub
