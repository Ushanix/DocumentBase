Option Explicit

' ============================================
' Module   : Pst_OutputToObsidian
' Layer    : Presentation
' Purpose  : Export Collection documents to Obsidian markdown
'            with YAML frontmatter
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Adapted from FlowBase Pst_OutputToObsidian
'
' Output structure:
'   <output_dir>/
'     README.md                           <- Collection header
'     <document_id>_<version>_<title>.md  <- Each document row
'
' Path resolution:
'   if HeaderInfo.collection_output_path is non-empty -> use as-is
'   else -> DEF_Parameter.OUTPUT_ROOT / collection_id
' ============================================

Private Const TOOL_NAME As String = "OutputToObsidian"

' Document columns used for frontmatter
Private Const COL_NO As String = "no"
Private Const COL_TITLE As String = "title"
Private Const COL_DOCUMENT_ID As String = "document_id"
Private Const COL_DOC_TYPE_PREFIX As String = "doc_type_prefix"

' ============================================
' OutputAll
' Export ALL DOC- sheets to Obsidian (batch)
' Called from UI_Dashboard action: output_all
' ============================================
Public Sub OutputAll()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputAll: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Exporting all collections to Obsidian..."
    Application.ScreenUpdating = False

    Dim docSheets As Collection
    Set docSheets = FilterSheetsByPrefix(PREFIX_COLLECTION)

    Dim successCount As Long
    Dim skipCount As Long
    Dim errorCount As Long
    successCount = 0
    skipCount = 0
    errorCount = 0

    Dim sheetName As Variant
    For Each sheetName In docSheets
        ' Skip template
        If Left(CStr(sheetName), Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
            skipCount = skipCount + 1
            GoTo NextAllSheet
        End If

        Application.StatusBar = "Exporting: " & CStr(sheetName) & "..."

        Dim result As String
        result = OutputCollectionSheet(CStr(sheetName))

        If Left(result, 5) = "ERROR" Then
            errorCount = errorCount + 1
            LogError TOOL_NAME, result
        Else
            successCount = successCount + 1
        End If

NextAllSheet:
    Next sheetName

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputAll: Completed"
    LogInfo TOOL_NAME, "  Success: " & successCount & ", Skip: " & skipCount & ", Error: " & errorCount
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Export all completed." & vbCrLf & vbCrLf & _
           "Success: " & successCount & " collections" & vbCrLf & _
           "Skipped: " & skipCount & vbCrLf & _
           "Errors: " & errorCount, vbInformation, "Complete"
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "OutputAll Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' OutputToObsidian
' Export current DOC- sheet to Obsidian (single)
' ============================================
Public Sub OutputToObsidian()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputToObsidian: Started"
    LogInfo TOOL_NAME, "========================================"

    Dim currentSheet As String
    currentSheet = ActiveSheet.Name

    If Left(currentSheet, Len(PREFIX_COLLECTION)) <> PREFIX_COLLECTION Then
        MsgBox "Please run from a DOC-* sheet.", vbExclamation, "Error"
        Exit Sub
    End If

    If Left(currentSheet, Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
        MsgBox "Cannot export template sheet.", vbExclamation, "Error"
        Exit Sub
    End If

    Application.StatusBar = "Exporting to Obsidian..."
    Application.ScreenUpdating = False

    Dim result As String
    result = OutputCollectionSheet(currentSheet)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputToObsidian: Completed"
    LogInfo TOOL_NAME, "========================================"

    If Left(result, 5) = "ERROR" Then
        MsgBox result, vbExclamation, "Error"
    Else
        MsgBox result, vbInformation, "Complete"
    End If

    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' OutputCollectionSheet
' Export a single Collection sheet
'
' Returns:
'   Result message (starts with "ERROR" on failure)
' ============================================
Private Function OutputCollectionSheet(sheetName As String) As String
    On Error GoTo EH

    LogInfo TOOL_NAME, "Processing: " & sheetName

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    ' --- Read HeaderInfo ---
    Dim headerInfo As Object
    Set headerInfo = ReadHeaderInfo(ws)
    If headerInfo.Count = 0 Then
        OutputCollectionSheet = "ERROR: Failed to read DOC_HeaderInfo"
        Exit Function
    End If

    Dim collectionId As String
    If headerInfo.Exists("collection_id") Then
        collectionId = CStr(headerInfo("collection_id"))
    Else
        collectionId = sheetName
    End If

    ' --- Resolve output directory ---
    Dim outputDir As String
    outputDir = ResolveOutputPath(headerInfo, collectionId)

    If Len(outputDir) = 0 Then
        OutputCollectionSheet = "ERROR: OUTPUT_ROOT not configured and output_path not set"
        Exit Function
    End If

    LogInfo TOOL_NAME, "Output dir: " & outputDir

    ' --- Create output folder ---
    If Not CreateFolder(outputDir) Then
        OutputCollectionSheet = "ERROR: Failed to create folder: " & outputDir
        Exit Function
    End If

    ' --- Export Collection README ---
    Dim readmeWritten As Boolean
    readmeWritten = WriteCollectionReadme(outputDir, headerInfo)

    ' --- Export individual documents ---
    Dim docCount As Long
    docCount = WriteDocumentFiles(ws, outputDir, headerInfo)

    ' --- Update HeaderInfo.collection_updated ---
    Dim headerMarkerRow As Long
    headerMarkerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
    If headerMarkerRow > 0 Then
        UpdateKeyValueTable ws, headerMarkerRow + 1, "collection_updated", Format(Date, "yyyy-mm-dd")
    End If

    LogInfo TOOL_NAME, "Completed: " & sheetName & " (README + " & docCount & " docs)"

    OutputCollectionSheet = "Export completed: " & sheetName & vbCrLf & _
                            "Output: " & outputDir & vbCrLf & _
                            "1 README + " & docCount & " documents."
    Exit Function

EH:
    LogError TOOL_NAME, "Error in " & sheetName & ": " & Err.Description
    OutputCollectionSheet = "ERROR: " & Err.Description
End Function

' ============================================
' ResolveOutputPath
' Resolve output directory using priority:
'   1. HeaderInfo.collection_output_path (if non-empty)
'   2. DEF_Parameter.OUTPUT_ROOT / <collection_id>_<collection_name>
'
' Folder name format: DOC-TECH-01_Git運用ガイド
' collection_name is sanitized for filesystem safety.
' ============================================
Private Function ResolveOutputPath(headerInfo As Object, collectionId As String) As String
    ResolveOutputPath = ""

    ' Priority 1: collection_output_path from HeaderInfo
    If headerInfo.Exists("collection_output_path") Then
        Dim customPath As String
        customPath = Trim(CStr(headerInfo("collection_output_path")))
        If Len(customPath) > 0 Then
            ResolveOutputPath = customPath
            LogInfo TOOL_NAME, "Using custom output_path: " & customPath
            Exit Function
        End If
    End If

    ' Priority 2: OUTPUT_ROOT / collection_id_collection_name
    Dim outputRoot As String
    outputRoot = GetOutputRoot()

    If Len(outputRoot) > 0 Then
        Dim folderName As String
        folderName = BuildCollectionFolderName(headerInfo, collectionId)
        ResolveOutputPath = BuildFilePath(outputRoot, folderName)
        LogInfo TOOL_NAME, "Using OUTPUT_ROOT: " & ResolveOutputPath
    End If
End Function

' ============================================
' BuildCollectionFolderName
' Generate folder name: <collection_id>_<collection_name>
' Falls back to collection_id only if name is empty.
' ============================================
Private Function BuildCollectionFolderName(headerInfo As Object, collectionId As String) As String
    Dim collName As String
    collName = ""

    If headerInfo.Exists("collection_name") Then
        collName = Trim(CStr(headerInfo("collection_name")))
    End If

    If Len(collName) > 0 Then
        BuildCollectionFolderName = collectionId & "_" & SanitizeFilename(collName)
    Else
        BuildCollectionFolderName = collectionId
    End If
End Function

' ============================================
' GetOutputRoot
' Get OUTPUT_ROOT from DEF_Parameter
' ============================================
Private Function GetOutputRoot() As String
    GetOutputRoot = ""

    If Not SheetExists(SHEET_DEF_PARAMETER) Then Exit Function

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)

    Dim result As Variant
    result = LookupTableValue(ws, TBL_DEF_PARAMETER, "name", "value", PARAM_OUTPUT_ROOT)

    If Not IsEmpty(result) And Len(CStr(result)) > 0 Then
        GetOutputRoot = CStr(result)
    End If
End Function

' ============================================
' ReadHeaderInfo
' Read DOC_HeaderInfo key-value table
' ============================================
Private Function ReadHeaderInfo(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
    If markerRow = 0 Then
        Set ReadHeaderInfo = dict
        Exit Function
    End If

    Set ReadHeaderInfo = ReadKeyValueTable(ws, markerRow + 1)
End Function

' ============================================
' WriteCollectionReadme
' Write Collection README.md with YAML frontmatter
'
' Frontmatter fields (from SHEET_DESIGN.md 7.2):
'   collection_id, collection_name, collection_domain,
'   collection_related_project, collection_status,
'   collection_created, collection_updated
' ============================================
Private Function WriteCollectionReadme(outputDir As String, headerInfo As Object) As Boolean
    WriteCollectionReadme = False

    Dim lines As Collection
    Set lines = New Collection

    ' --- YAML frontmatter ---
    lines.Add "---"

    AppendYaml lines, "collection_id", headerInfo
    AppendYaml lines, "collection_name", headerInfo
    AppendYaml lines, "collection_domain", headerInfo
    AppendYaml lines, "collection_related_project", headerInfo
    AppendYaml lines, "collection_status", headerInfo
    AppendYaml lines, "collection_created", headerInfo
    AppendYaml lines, "collection_updated", headerInfo

    lines.Add "---"
    lines.Add ""

    ' --- Body ---
    Dim title As String
    If headerInfo.Exists("collection_name") Then
        title = CStr(headerInfo("collection_name"))
    End If

    If Len(title) > 0 Then
        lines.Add "# " & title
        lines.Add ""
    End If

    If headerInfo.Exists("collection_summary") Then
        Dim summary As String
        summary = CStr(headerInfo("collection_summary"))
        If Len(summary) > 0 Then
            lines.Add summary
            lines.Add ""
        End If
    End If

    ' Build content
    Dim content As String
    content = JoinLines(lines)

    ' Write or update file
    Dim filepath As String
    filepath = BuildFilePath(outputDir, "README.md")

    If FileExists(filepath) Then
        Dim existing As String
        existing = ReadTextFile(filepath)
        Dim body As String
        body = ExtractBodyAfterFrontmatter(existing)
        ' Rebuild: new frontmatter + existing body
        Dim fmOnly As String
        fmOnly = ExtractFrontmatter(lines)
        content = fmOnly & body
    End If

    WriteCollectionReadme = WriteTextFile(filepath, content, False)

    If WriteCollectionReadme Then
        LogInfo TOOL_NAME, "Written: " & filepath
    Else
        LogError TOOL_NAME, "Failed to write: " & filepath
    End If
End Function

' ============================================
' WriteDocumentFiles
' Write individual document .md files
'
' Filename: <document_id>_<version>_<title>.md
' Frontmatter: all DOC_DocumentList columns except no
'
' Returns:
'   Number of files written
' ============================================
Private Function WriteDocumentFiles(ws As Worksheet, outputDir As String, headerInfo As Object) As Long
    WriteDocumentFiles = 0

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_DOCUMENT_LIST)
    If markerRow = 0 Then Exit Function

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim headers As Variant
    headers = GetTableHeaders(ws, headerRow)

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ' Find key column indices
    Dim colTitle As Long: colTitle = GetColumnIndex(headers, COL_TITLE)
    Dim colDocId As Long: colDocId = GetColumnIndex(headers, COL_DOCUMENT_ID)
    Dim colNo As Long: colNo = GetColumnIndex(headers, COL_NO)

    Dim written As Long
    written = 0

    Dim r As Long
    For r = headerRow + 1 To headerRow + 300
        ' Check if row has data
        Dim noVal As Variant
        If colNo > 0 Then
            noVal = ws.Cells(r, colNo).Value
        Else
            noVal = ws.Cells(r, 1).Value
        End If
        If IsEmpty(noVal) Then Exit For

        ' Read row values into dictionary
        Dim rowData As Object
        Set rowData = CreateObject("Scripting.Dictionary")

        Dim j As Long
        Dim lb As Long: lb = LBound(headers)
        For j = lb To UBound(headers)
            Dim colIdx As Long
            colIdx = j - lb + 1
            Dim cellVal As Variant
            cellVal = ws.Cells(r, colIdx).Value
            If Not IsEmpty(cellVal) Then
                rowData(headers(j)) = cellVal
            End If
        Next j

        ' Need title and document_id
        Dim title As String: title = ""
        Dim docId As String: docId = ""
        Dim version As String: version = ""

        If rowData.Exists("title") Then title = CStr(rowData("title"))
        If rowData.Exists("document_id") Then docId = CStr(rowData("document_id"))
        If rowData.Exists("version") Then version = CStr(rowData("version"))

        If Len(title) = 0 Then GoTo NextDoc

        ' --- Generate filename ---
        Dim filenameBase As String
        If Len(docId) > 0 Then
            If Len(version) > 0 Then
                filenameBase = docId & "_" & version & "_" & title
            Else
                filenameBase = docId & "_" & title
            End If
        Else
            filenameBase = title
        End If

        Dim filename As String
        filename = SanitizeFilename(filenameBase) & ".md"

        Dim filepath As String
        filepath = BuildFilePath(outputDir, filename)

        ' --- Generate YAML frontmatter ---
        Dim fmLines As Collection
        Set fmLines = New Collection

        fmLines.Add "---"

        ' Output all columns except 'no' in header order
        For j = lb To UBound(headers)
            Dim hdr As String
            hdr = headers(j)

            ' Skip 'no' (internal row number) and 'doc_type_prefix' (internal derived value)
            If hdr = COL_NO Or hdr = COL_DOC_TYPE_PREFIX Then GoTo NextCol

            If rowData.Exists(hdr) Then
                Dim yamlVal As String
                yamlVal = FormatYamlValue(rowData(hdr))
                If Len(yamlVal) > 0 Then
                    fmLines.Add hdr & ": " & yamlVal
                End If
            End If
NextCol:
        Next j

        fmLines.Add "---"
        fmLines.Add ""

        ' --- Build content ---
        Dim content As String

        If FileExists(filepath) Then
            ' Preserve existing body
            Dim existing As String
            existing = ReadTextFile(filepath)
            Dim body As String
            body = ExtractBodyAfterFrontmatter(existing)
            content = JoinLines(fmLines) & body
        Else
            ' New file: frontmatter + heading
            content = JoinLines(fmLines)
            content = content & "# " & title & vbLf & vbLf

            ' Add summary if available
            If rowData.Exists("summary") Then
                Dim summ As String
                summ = CStr(rowData("summary"))
                If Len(summ) > 0 Then
                    content = content & summ & vbLf
                End If
            End If
        End If

        ' --- Write file ---
        If WriteTextFile(filepath, content, False) Then
            written = written + 1
            LogDebug TOOL_NAME, "  Written: " & filename
        Else
            LogError TOOL_NAME, "  Failed: " & filename
        End If

NextDoc:
    Next r

    WriteDocumentFiles = written
End Function

' ============================================
' Helper functions
' ============================================

Private Sub AppendYaml(lines As Collection, key As String, dict As Object)
    If dict.Exists(key) Then
        Dim val As String
        val = FormatYamlValue(dict(key))
        If Len(val) > 0 Then
            lines.Add key & ": " & val
        End If
    End If
End Sub

Private Function FormatYamlValue(val As Variant) As String
    FormatYamlValue = ""

    If IsEmpty(val) Or IsNull(val) Then Exit Function

    If IsDate(val) Then
        FormatYamlValue = Format(val, "yyyy-mm-dd")
        Exit Function
    End If

    If VarType(val) = vbBoolean Then
        FormatYamlValue = IIf(val, "true", "false")
        Exit Function
    End If

    If IsNumeric(val) And VarType(val) <> vbString Then
        FormatYamlValue = CStr(val)
        Exit Function
    End If

    Dim s As String
    s = CStr(val)
    If Left(s, 1) = "=" Then Exit Function  ' skip formulas
    If Len(Trim(s)) = 0 Then Exit Function

    ' Quote if contains special chars
    If InStr(s, ":") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, """") > 0 Then
        s = Replace(s, """", "\""")
        FormatYamlValue = """" & s & """"
    Else
        FormatYamlValue = s
    End If
End Function

Private Function ExtractBodyAfterFrontmatter(content As String) As String
    Dim lines() As String
    lines = Split(content, vbLf)

    If UBound(lines) < 0 Then
        ExtractBodyAfterFrontmatter = content
        Exit Function
    End If

    If Trim(lines(0)) <> "---" Then
        ExtractBodyAfterFrontmatter = content
        Exit Function
    End If

    Dim i As Long
    For i = 1 To UBound(lines)
        If Trim(lines(i)) = "---" Then
            Dim result As String
            Dim j As Long
            For j = i + 1 To UBound(lines)
                If j = i + 1 Then
                    result = lines(j)
                Else
                    result = result & vbLf & lines(j)
                End If
            Next j
            ExtractBodyAfterFrontmatter = result
            Exit Function
        End If
    Next i

    ExtractBodyAfterFrontmatter = content
End Function

Private Function ExtractFrontmatter(lines As Collection) As String
    Dim result As String
    Dim line As Variant
    Dim pastSecondDash As Boolean
    pastSecondDash = False

    Dim dashCount As Long
    dashCount = 0

    For Each line In lines
        result = result & line & vbLf
        If CStr(line) = "---" Then
            dashCount = dashCount + 1
            If dashCount = 2 Then
                result = result & vbLf
                ExtractFrontmatter = result
                Exit Function
            End If
        End If
    Next line

    ExtractFrontmatter = result
End Function

Private Function JoinLines(lines As Collection) As String
    Dim result As String
    Dim line As Variant
    For Each line In lines
        result = result & line & vbLf
    Next line
    JoinLines = result
End Function
