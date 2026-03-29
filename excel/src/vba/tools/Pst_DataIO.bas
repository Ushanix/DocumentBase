Attribute VB_Name = "Pst_DataIO"
Option Explicit

' ============================================
' Module   : Pst_DataIO
' Layer    : Presentation
' Purpose  : DataIO -- RefreshList, ExportSelected, ImportSelected
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Adapted from FlowBase Pst_DataIO for DocumentBase
'            Handles DOC-/DEF_/TPL_ sheets
' ============================================

Private Const TOOL_NAME As String = "DataIO"

' ============================================
' RefreshList
' Scan DOC-/DEF_/TPL_ sheets -> ExportList
' Scan data_path/*.yaml -> ImportList
' ============================================
Public Sub RefreshList()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "RefreshList: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Refreshing DataIO lists..."
    Application.ScreenUpdating = False

    If Not SheetExists(SHEET_UI_DATA_IO) Then
        MsgBox SHEET_UI_DATA_IO & " not found.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim wsIO As Worksheet
    Set wsIO = ThisWorkbook.Worksheets(SHEET_UI_DATA_IO)

    Dim dataPath As String
    dataPath = GetDataPath(wsIO)

    If Len(dataPath) = 0 Then
        MsgBox "data_path is not configured.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "data_path = " & dataPath

    ' ExportList
    Dim exportMkr As Long
    exportMkr = FindTblStartRow(wsIO, TBL_EXPORT_LIST)
    If exportMkr > 0 Then
        PopulateExportList wsIO, exportMkr + 1
    End If

    ' ImportList
    Dim importMkr As Long
    importMkr = FindTblStartRow(wsIO, TBL_IMPORT_LIST)
    If importMkr > 0 Then
        PopulateImportList wsIO, importMkr + 1, dataPath
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "RefreshList: Completed"
    MsgBox "List refreshed.", vbInformation, "Complete"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "RefreshList Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' ExportSelected
' Export sheets marked select=YES to YAML files
' ============================================
Public Sub ExportSelected()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "ExportSelected: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Exporting..."
    Application.ScreenUpdating = False

    Dim wsIO As Worksheet
    Set wsIO = ThisWorkbook.Worksheets(SHEET_UI_DATA_IO)

    Dim dataPath As String
    dataPath = GetDataPath(wsIO)
    If Len(dataPath) = 0 Then
        MsgBox "data_path is not configured.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    If Not CreateFolder(dataPath) Then
        MsgBox "Cannot create: " & dataPath, vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Read ExportList
    Dim exportMkr As Long
    exportMkr = FindTblStartRow(wsIO, TBL_EXPORT_LIST)
    If exportMkr = 0 Then
        MsgBox "ExportList not found.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim headerRow As Long
    headerRow = exportMkr + 1

    Dim tableData As Variant
    tableData = ReadTableData(wsIO, headerRow)

    Dim headers As Variant
    headers = tableData(0)

    Dim rows As Collection
    Set rows = tableData(1)

    Dim exportCount As Long
    exportCount = 0

    Dim archiveSheets As Collection
    Set archiveSheets = New Collection

    Dim row As Object
    For Each row In rows
        Dim selectVal As String
        selectVal = ""
        If row.Exists("select") Then selectVal = UCase(Trim(CStr(row("select"))))
        If selectVal <> "YES" Then GoTo NextExport

        Dim sheetName As String
        sheetName = CStr(row("sheet_name"))
        If Not SheetExists(sheetName) Then GoTo NextExport

        Dim sheetType As String
        sheetType = ""
        If row.Exists("type") Then sheetType = CStr(row("type"))
        If Len(sheetType) = 0 Then sheetType = DetermineSheetType(sheetName)

        LogInfo TOOL_NAME, "Exporting: " & sheetName

        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(sheetName)

        ' Build meta + tables
        Dim metaDict As Object
        Set metaDict = BuildMetaDict(sheetName, sheetType)

        Dim tablesDict As Object
        Set tablesDict = BuildExportData(ws)

        If tablesDict Is Nothing Then GoTo NextExport

        ' Serialize and write
        Dim yamlContent As String
        yamlContent = SerializeToYaml(metaDict, tablesDict)

        Dim fileName As String
        fileName = SanitizeFilename(sheetName) & ".yaml"

        If WriteTextFile(BuildFilePath(dataPath, fileName), yamlContent) Then
            exportCount = exportCount + 1

            Dim postAction As String
            postAction = "backup"
            If row.Exists("post_action") Then postAction = LCase(Trim(CStr(row("post_action"))))
            If postAction = "archive" Then
                archiveSheets.Add Array(sheetName, sheetType)
            End If
        End If

NextExport:
    Next row

    ' Process archives
    Dim item As Variant
    For Each item In archiveSheets
        Dim archName As String
        archName = item(0)
        Dim archType As String
        archType = UCase(CStr(item(1)))

        ' Guard: DEF/TPL sheets cannot be archived
        If archType = "DEF" Or archType = "TPL" Then
            LogWarn TOOL_NAME, "Cannot archive system sheet: " & archName
        Else
            Dim answer As VbMsgBoxResult
            answer = MsgBox("Archive (delete) sheet: " & archName & "?", vbYesNo + vbQuestion, "Confirm")
            If answer = vbYes Then
                Application.DisplayAlerts = False
                ThisWorkbook.Worksheets(archName).Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next item

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "ExportSelected: " & exportCount & " exported"
    MsgBox exportCount & " sheet(s) exported.", vbInformation, "Complete"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "ExportSelected Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' ImportSelected
' Import YAML files marked select=YES
' ============================================
Public Sub ImportSelected()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "ImportSelected: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Importing..."
    Application.ScreenUpdating = False

    Dim wsIO As Worksheet
    Set wsIO = ThisWorkbook.Worksheets(SHEET_UI_DATA_IO)

    Dim dataPath As String
    dataPath = GetDataPath(wsIO)
    If Len(dataPath) = 0 Then
        MsgBox "data_path is not configured.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim importMkr As Long
    importMkr = FindTblStartRow(wsIO, TBL_IMPORT_LIST)
    If importMkr = 0 Then
        MsgBox "ImportList not found.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim headerRow As Long
    headerRow = importMkr + 1

    Dim tableData As Variant
    tableData = ReadTableData(wsIO, headerRow)

    Dim rows As Collection
    Set rows = tableData(1)

    Dim importCount As Long
    importCount = 0

    Dim row As Object
    For Each row In rows
        Dim selectVal As String
        selectVal = ""
        If row.Exists("select") Then selectVal = UCase(Trim(CStr(row("select"))))
        If selectVal <> "YES" Then GoTo NextImport

        Dim action As String
        action = ""
        If row.Exists("action") Then action = LCase(Trim(CStr(row("action"))))
        If action = "skip" Or Len(action) = 0 Then GoTo NextImport

        Dim fileName As String
        fileName = ""
        If row.Exists("file_name") Then fileName = CStr(row("file_name"))
        If Len(fileName) = 0 Then GoTo NextImport

        Dim filePath As String
        filePath = BuildFilePath(dataPath, fileName)
        If Not FileExists(filePath) Then GoTo NextImport

        Dim sheetName As String
        sheetName = ""
        If row.Exists("sheet_name") Then sheetName = CStr(row("sheet_name"))

        ' Parse YAML
        Dim yamlData As Object
        Set yamlData = ParseYamlFile(filePath)
        If yamlData Is Nothing Then GoTo NextImport

        Dim meta As Object
        Set meta = yamlData("meta")

        If Len(sheetName) = 0 And Not meta Is Nothing Then
            If meta.Exists("sheet_name") Then sheetName = CStr(meta("sheet_name"))
        End If
        If Len(sheetName) = 0 Then GoTo NextImport

        LogInfo TOOL_NAME, "Importing: " & fileName & " -> " & sheetName & " (" & action & ")"

        Dim targetWs As Worksheet

        If action = "create" Then
            ' Determine template
            Dim templateName As String
            templateName = ""
            If Not meta Is Nothing And meta.Exists("template") Then
                templateName = CStr(meta("template"))
            End If
            If Len(templateName) = 0 Then
                templateName = GetTemplateForType(DetermineSheetType(sheetName))
            End If
            If Len(templateName) = 0 Or Not SheetExists(templateName) Then GoTo NextImport

            If SheetExists(sheetName) Then GoTo NextImport

            Set targetWs = CopySheet(templateName, sheetName)
            If targetWs Is Nothing Then GoTo NextImport

        ElseIf action = "overwrite" Then
            If Not SheetExists(sheetName) Then GoTo NextImport
            Set targetWs = ThisWorkbook.Worksheets(sheetName)
        Else
            GoTo NextImport
        End If

        ' Apply data
        If ApplyImportData(targetWs, yamlData) Then
            importCount = importCount + 1
        End If

NextImport:
    Next row

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "ImportSelected: " & importCount & " imported"
    MsgBox importCount & " file(s) imported.", vbInformation, "Complete"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "ImportSelected Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' Private helpers
' ============================================

Private Function GetDataPath(wsIO As Worksheet) As String
    GetDataPath = ""

    ' Try DataIOConfig
    Dim mkr As Long
    mkr = FindTblStartRow(wsIO, TBL_DATA_IO_CONFIG)
    If mkr > 0 Then
        Dim config As Object
        Set config = ReadKeyValueTable(wsIO, mkr + 1)
        Dim key As Variant
        For Each key In config.Keys
            If StrComp(CStr(key), "data_path", vbTextCompare) = 0 Then
                Dim p As String
                p = Trim(CStr(config(key)))
                If Len(p) > 0 Then
                    GetDataPath = p
                    Exit Function
                End If
            End If
        Next key
    End If

    ' Fallback: DEF_Parameter
    If SheetExists(SHEET_DEF_PARAMETER) Then
        Dim wsDef As Worksheet
        Set wsDef = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)
        Dim val As Variant
        val = LookupTableValue(wsDef, TBL_DEF_PARAMETER, "name", "value", PARAM_DATA_EXPORT_PATH)
        If Not IsEmpty(val) And Len(CStr(val)) > 0 Then
            GetDataPath = CStr(val)
        End If
    End If
End Function

Private Sub PopulateExportList(wsIO As Worksheet, headerRow As Long)
    Dim headers As Variant
    headers = GetTableHeaders(wsIO, headerRow)

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ClearTableData wsIO, headerRow, colCount

    Dim rowNum As Long
    rowNum = headerRow + 1

    Dim seqNo As Long
    seqNo = 0

    ' DOC- sheets
    Dim docSheets As Collection
    Set docSheets = FilterSheetsByPrefix(PREFIX_COLLECTION)

    Dim sheetName As Variant
    For Each sheetName In docSheets
        seqNo = seqNo + 1
        Dim rowDict As Object
        Set rowDict = BuildExportListRow(CStr(sheetName), "DOC", seqNo)
        WriteTableRow wsIO, rowNum, headers, rowDict, "sheet_name"
        rowNum = rowNum + 1
    Next sheetName

    ' DEF_ sheets (exclude templates)
    Dim defSheets As Collection
    Set defSheets = FilterSheetsByPrefix(PREFIX_DEFINITION)

    For Each sheetName In defSheets
        seqNo = seqNo + 1
        Set rowDict = BuildExportListRow(CStr(sheetName), "DEF", seqNo)
        WriteTableRow wsIO, rowNum, headers, rowDict, "sheet_name"
        rowNum = rowNum + 1
    Next sheetName

    ' TPL_ sheets
    Dim tplSheets As Collection
    Set tplSheets = FilterSheetsByPrefix(PREFIX_TEMPLATE)

    For Each sheetName In tplSheets
        seqNo = seqNo + 1
        Set rowDict = BuildExportListRow(CStr(sheetName), "TPL", seqNo)
        WriteTableRow wsIO, rowNum, headers, rowDict, "sheet_name"
        rowNum = rowNum + 1
    Next sheetName

    ' Apply validations
    ApplyColumnValidation wsIO, headerRow, headers, "select", "YES", seqNo
    ApplyColumnValidation wsIO, headerRow, headers, "post_action", "backup,archive", seqNo

    LogInfo TOOL_NAME, "ExportList: " & seqNo & " rows"
End Sub

Private Function BuildExportListRow(sheetName As String, sheetType As String, seqNo As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d("select") = ""
    d("no") = seqNo
    d("sheet_name") = sheetName
    d("type") = sheetType
    d("post_action") = "backup"
    d("collection_name") = ""
    d("status") = ""
    d("last_update") = ""

    ' Try DOC_HeaderInfo
    If SheetExists(sheetName) Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(sheetName)

        Dim mkr As Long
        mkr = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
        If mkr > 0 Then
            Dim kv As Object
            Set kv = ReadKeyValueTable(ws, mkr + 1)
            If kv.Exists("collection_name") Then d("collection_name") = kv("collection_name")
            If kv.Exists("status") Then d("status") = kv("status")
            If kv.Exists("updated") Then d("last_update") = kv("updated")
        End If
    End If

    Set BuildExportListRow = d
End Function

Private Sub PopulateImportList(wsIO As Worksheet, headerRow As Long, dataPath As String)
    Dim headers As Variant
    headers = GetTableHeaders(wsIO, headerRow)

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ClearTableData wsIO, headerRow, colCount

    Dim yamlFiles As Collection
    Set yamlFiles = ScanYamlFiles(dataPath)

    If yamlFiles.Count = 0 Then
        LogInfo TOOL_NAME, "No YAML files in: " & dataPath
        Exit Sub
    End If

    Dim rowNum As Long
    rowNum = headerRow + 1

    Dim seqNo As Long
    seqNo = 0

    Dim fName As Variant
    For Each fName In yamlFiles
        Dim filePath As String
        filePath = BuildFilePath(dataPath, CStr(fName))

        Dim meta As Object
        Set meta = ParseYamlMeta(filePath)

        seqNo = seqNo + 1

        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("select") = ""
        d("no") = seqNo
        d("file_name") = CStr(fName)

        Dim yamlSheet As String
        yamlSheet = ""
        If Not meta Is Nothing Then
            If meta.Exists("sheet_name") Then yamlSheet = CStr(meta("sheet_name"))
            d("sheet_name") = yamlSheet
            If meta.Exists("exported_at") Then
                d("exported_at") = meta("exported_at")
            Else
                d("exported_at") = ""
            End If
        Else
            d("sheet_name") = ""
            d("exported_at") = ""
        End If

        Dim exists As Boolean
        exists = False
        If Len(yamlSheet) > 0 Then exists = SheetExists(yamlSheet)

        d("exists") = IIf(exists, "YES", "NO")
        d("action") = IIf(exists, "skip", "create")

        WriteTableRow wsIO, rowNum, headers, d
        rowNum = rowNum + 1
    Next fName

    ApplyColumnValidation wsIO, headerRow, headers, "select", "YES", seqNo
    ApplyColumnValidation wsIO, headerRow, headers, "action", "create,overwrite,skip", seqNo

    LogInfo TOOL_NAME, "ImportList: " & seqNo & " rows"
End Sub

Private Function BuildExportData(ws As Worksheet) As Object
    Dim tablesDict As Object
    Set tablesDict = CreateObject("Scripting.Dictionary")

    Dim allMarkers As Object
    Set allMarkers = FindAllTblMarkers(ws)

    ' Skip list -- UI/action/computed tables
    Dim skipDict As Object
    Set skipDict = CreateObject("Scripting.Dictionary")
    Dim skipArr As Variant
    skipArr = GetExportSkipTables()
    Dim s As Long
    For s = LBound(skipArr) To UBound(skipArr)
        skipDict(CStr(skipArr(s))) = True
    Next s

    ' KV table list
    Dim kvDict As Object
    Set kvDict = CreateObject("Scripting.Dictionary")
    Dim kvArr As Variant
    kvArr = GetKeyValueTables()
    Dim k As Long
    For k = LBound(kvArr) To UBound(kvArr)
        kvDict(CStr(kvArr(k))) = True
    Next k

    Dim markerName As Variant
    For Each markerName In allMarkers.Keys
        If skipDict.Exists(CStr(markerName)) Then GoTo NextMarker

        Dim markerRow As Long
        markerRow = allMarkers(markerName)

        If kvDict.Exists(CStr(markerName)) Then
            Dim kvData As Object
            Set kvData = ReadKeyValueTable(ws, markerRow + 1)
            Dim kvInfo As Object
            Set kvInfo = CreateObject("Scripting.Dictionary")
            kvInfo("type") = "key_value"
            Set kvInfo("data") = kvData
            Set tablesDict(CStr(markerName)) = kvInfo
        Else
            Dim tblData As Variant
            tblData = ReadTableData(ws, markerRow + 1)
            Dim tblInfo As Object
            Set tblInfo = CreateObject("Scripting.Dictionary")
            tblInfo("type") = "tabular"
            tblInfo("headers") = tblData(0)
            Set tblInfo("rows") = tblData(1)
            Set tablesDict(CStr(markerName)) = tblInfo
        End If

NextMarker:
    Next markerName

    If tablesDict.Count = 0 Then
        Set BuildExportData = Nothing
    Else
        Set BuildExportData = tablesDict
    End If
End Function

Private Function BuildMetaDict(sheetName As String, sheetType As String) As Object
    Dim meta As Object
    Set meta = CreateObject("Scripting.Dictionary")
    meta("schema_version") = "1.0"
    meta("exported_at") = Format(Now, "yyyy-mm-dd\Thh:nn:ss")
    meta("source_version") = "0.1.1"
    meta("sheet_name") = sheetName
    meta("sheet_type") = UCase(sheetType)
    meta("template") = GetTemplateForType(sheetType)
    Set BuildMetaDict = meta
End Function

Private Function GetTemplateForType(sheetType As String) As String
    GetTemplateForType = ""
    Dim upperType As String
    upperType = UCase(sheetType)
    If upperType = "DOC" Then
        GetTemplateForType = DEFAULT_COLLECTION_TEMPLATE
    End If
End Function

Private Function ApplyImportData(ws As Worksheet, yamlData As Object) As Boolean
    On Error GoTo EH
    ApplyImportData = False

    Dim tablesDict As Object
    Set tablesDict = yamlData("tables")
    If tablesDict Is Nothing Then Exit Function

    Dim tblName As Variant
    For Each tblName In tablesDict.Keys
        Dim tblInfo As Object
        Set tblInfo = tablesDict(tblName)
        Dim tblType As String
        tblType = CStr(tblInfo("type"))

        If tblType = "key_value" Then
            WriteKeyValueData ws, CStr(tblName), tblInfo("data")
        ElseIf tblType = "tabular" Then
            WriteTabularData ws, CStr(tblName), tblInfo("headers"), tblInfo("rows")
        End If
    Next tblName

    ApplyImportData = True
    Exit Function

EH:
    LogError TOOL_NAME, "ApplyImportData Error: " & Err.Description
    ApplyImportData = False
End Function

Private Function WriteKeyValueData(ws As Worksheet, markerName As String, kvData As Object) As Long
    WriteKeyValueData = 0
    Dim mkr As Long
    mkr = FindTblStartRow(ws, markerName)
    If mkr = 0 Then Exit Function

    ClearKeyValueTableValues ws, mkr + 1

    Dim key As Variant
    For Each key In kvData.Keys
        If UpdateKeyValueTable(ws, mkr + 1, CStr(key), kvData(key)) Then
            WriteKeyValueData = WriteKeyValueData + 1
        End If
    Next key
End Function

Private Function WriteTabularData(ws As Worksheet, markerName As String, _
                                   yamlHeaders As Variant, yamlRows As Collection) As Long
    WriteTabularData = 0
    Dim mkr As Long
    mkr = FindTblStartRow(ws, markerName)
    If mkr = 0 Then Exit Function

    Dim headerRow As Long
    headerRow = mkr + 1

    Dim sheetHeaders As Variant
    sheetHeaders = GetTableHeaders(ws, headerRow)

    Dim colCount As Long
    colCount = UBound(sheetHeaders) - LBound(sheetHeaders) + 1

    ClearTableData ws, headerRow, colCount

    Dim rowNum As Long
    rowNum = headerRow + 1

    Dim row As Object
    For Each row In yamlRows
        WriteTableRow ws, rowNum, sheetHeaders, row
        rowNum = rowNum + 1
        WriteTabularData = WriteTabularData + 1
    Next row
End Function

Private Function ClearKeyValueTableValues(ws As Worksheet, headerRow As Long, _
                                           Optional maxRows As Long = 100) As Long
    Dim i As Long
    Dim cleared As Long
    cleared = 0
    For i = headerRow + 1 To headerRow + maxRows
        If IsEmpty(ws.Cells(i, 1).Value) Or Trim(CStr(ws.Cells(i, 1).Value)) = "" Then Exit For
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            ws.Cells(i, 2).Value = Empty
            cleared = cleared + 1
        End If
    Next i
    ClearKeyValueTableValues = cleared
End Function

Private Function DetermineSheetType(sheetName As String) As String
    If Left(sheetName, Len(PREFIX_COLLECTION)) = PREFIX_COLLECTION Then
        DetermineSheetType = "DOC"
    ElseIf Left(sheetName, Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
        DetermineSheetType = "TPL"
    ElseIf Left(sheetName, Len(PREFIX_DEFINITION)) = PREFIX_DEFINITION Then
        DetermineSheetType = "DEF"
    ElseIf Left(sheetName, Len(PREFIX_UI)) = PREFIX_UI Then
        DetermineSheetType = "UI"
    ElseIf Left(sheetName, Len(PREFIX_LOG)) = PREFIX_LOG Then
        DetermineSheetType = "LOG"
    Else
        DetermineSheetType = ""
    End If
End Function

Private Function ScanYamlFiles(folderPath As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Set ScanYamlFiles = result
        Exit Function
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    Dim file As Object
    For Each file In folder.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "yaml" Or ext = "yml" Then
            result.Add file.Name
        End If
    Next file

    Set ScanYamlFiles = result
End Function

Private Sub ApplyColumnValidation(ws As Worksheet, headerRow As Long, headers As Variant, _
                                   colName As String, formula1 As String, dataRowCount As Long)
    If dataRowCount = 0 Then Exit Sub

    Dim colIdx As Long
    colIdx = GetColumnIndex(headers, colName)
    If colIdx = 0 Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(headerRow + 1, colIdx), ws.Cells(headerRow + dataRowCount, colIdx))

    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0

    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, formula1:=formula1
    rng.Validation.InCellDropdown = True
End Sub

' ============================================
' Export skip tables -- tables NOT exported to YAML
' ============================================
Private Function GetExportSkipTables() As Variant
    GetExportSkipTables = Array( _
        "UI_Operations", "UI_Status", "UI_SheetIndex", _
        "Actions", "IndexHeader", "AddCollection", _
        "DataIOConfig", "ExportList", "ImportList", _
        "DOC_Action", "LOG_UpdateHistory")
End Function
