Attribute VB_Name = "Pst_UpdateIndex"
Option Explicit

' ============================================
' Module   : Pst_UpdateIndex
' Layer    : Presentation
' Purpose  : Update UI_Dashboard (SheetIndex + Status) and
'            UI_CollectionIndex with collection info
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Combines FlowBase's IndexUpdate + ProjectIndexUpdate
'            into a single UpdateIndex for DocumentBase
' ============================================

Private Const TOOL_NAME As String = "UpdateIndex"

' Special columns
Private Const COL_NO As String = "no"
Private Const COL_SHEET_NAME As String = "sheet_name"

' ============================================
' UpdateIndex (Public entry point)
' Updates both UI_Dashboard and UI_CollectionIndex
' ============================================
Public Sub UpdateIndex()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateIndex: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Updating Index..."
    Application.ScreenUpdating = False

    ' --- Part 1: UI_Dashboard Tbl:UI_SheetIndex ---
    UpdateSheetIndex

    ' --- Part 2: UI_Dashboard Tbl:UI_Status ---
    UpdateStatus

    ' --- Part 3: UI_CollectionIndex Tbl:CollectionIndex ---
    UpdateCollectionIndex

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateIndex: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Index update completed.", vbInformation, "Complete"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' UpdateSheetIndex
' Collect all sheet info and write to UI_Dashboard Tbl:UI_SheetIndex
' ============================================
Private Sub UpdateSheetIndex()
    If Not SheetExists(SHEET_UI_DASHBOARD) Then
        LogWarn TOOL_NAME, "Sheet not found: " & SHEET_UI_DASHBOARD
        Exit Sub
    End If

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets(SHEET_UI_DASHBOARD)

    ' Find UI_SheetIndex marker
    Dim markerRow As Long
    markerRow = FindTblStartRow(wsDash, TBL_UI_SHEET_INDEX)
    If markerRow = 0 Then
        LogWarn TOOL_NAME, TBL_MARKER_PREFIX & TBL_UI_SHEET_INDEX & " not found"
        Exit Sub
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim headers As Variant
    headers = GetTableHeaders(wsDash, headerRow)

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ' Collect all sheets
    Dim sheetsData As Collection
    Set sheetsData = CollectAllSheets(headers)

    ' Sort by sheet_name
    Dim sortedData As Collection
    Set sortedData = SortByField(sheetsData, COL_SHEET_NAME)

    ' Clear + write
    ClearTableData wsDash, headerRow, colCount

    Dim dataRow As Long
    dataRow = headerRow + 1

    Dim rowNum As Long
    rowNum = 0

    Dim item As Object
    For Each item In sortedData
        rowNum = rowNum + 1
        item(COL_NO) = rowNum
        WriteTableRow wsDash, dataRow, headers, item, COL_SHEET_NAME
        dataRow = dataRow + 1
    Next item

    LogInfo TOOL_NAME, "SheetIndex: " & rowNum & " sheets"
End Sub

' ============================================
' UpdateStatus
' Update UI_Dashboard Tbl:UI_Status with summary metrics
' ============================================
Private Sub UpdateStatus()
    If Not SheetExists(SHEET_UI_DASHBOARD) Then Exit Sub

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Worksheets(SHEET_UI_DASHBOARD)

    Dim markerRow As Long
    markerRow = FindTblStartRow(wsDash, TBL_UI_STATUS)
    If markerRow = 0 Then
        LogWarn TOOL_NAME, TBL_MARKER_PREFIX & TBL_UI_STATUS & " not found"
        Exit Sub
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    ' Count collections
    Dim collectionSheets As Collection
    Set collectionSheets = FilterSheetsByPrefix(PREFIX_COLLECTION)

    ' Filter out templates (TPL_ prefix sheets that start with DOC- after copy)
    Dim totalCollections As Long
    totalCollections = 0
    Dim sheetName As Variant
    For Each sheetName In collectionSheets
        ' Skip if it's the template
        If CStr(sheetName) <> SHEET_TPL_COLLECTION Then
            totalCollections = totalCollections + 1
        End If
    Next sheetName

    ' Count total documents and active collections
    Dim totalDocuments As Long
    totalDocuments = 0
    Dim activeCollections As Long
    activeCollections = 0

    For Each sheetName In collectionSheets
        If CStr(sheetName) = SHEET_TPL_COLLECTION Then GoTo NextCollSheet

        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))

        ' Count documents in DOC_DocumentList
        Dim docMarkerRow As Long
        docMarkerRow = FindTblStartRow(ws, TBL_DOC_DOCUMENT_LIST)
        If docMarkerRow > 0 Then
            Dim docHeaderRow As Long
            docHeaderRow = docMarkerRow + 1
            Dim r As Long
            For r = docHeaderRow + 1 To docHeaderRow + 300
                If IsEmpty(ws.Cells(r, 1).Value) Then Exit For
                totalDocuments = totalDocuments + 1
            Next r
        End If

        ' Check status from DOC_HeaderInfo
        Dim headerMarkerRow As Long
        headerMarkerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
        If headerMarkerRow > 0 Then
            Dim headerInfo As Object
            Set headerInfo = ReadKeyValueTable(ws, headerMarkerRow + 1)
            If headerInfo.Exists("collection_status") Then
                If CStr(headerInfo("collection_status")) = "active" Then
                    activeCollections = activeCollections + 1
                End If
            End If
        End If

NextCollSheet:
    Next sheetName

    ' Write metrics
    UpdateKeyValueTable wsDash, headerRow, "total_collections", totalCollections
    UpdateKeyValueTable wsDash, headerRow, "total_documents", totalDocuments
    UpdateKeyValueTable wsDash, headerRow, "active_collections", activeCollections
    UpdateKeyValueTable wsDash, headerRow, "last_updated", Format(Now, "yyyy-mm-dd hh:mm")

    LogInfo TOOL_NAME, "Status: " & totalCollections & " collections, " & _
            totalDocuments & " documents, " & activeCollections & " active"
End Sub

' ============================================
' CollectionIndexUpdate (Public entry point)
' Standalone update of UI_CollectionIndex only.
' Called from UI_CollectionIndex Tbl:Actions button.
' ============================================
Public Sub CollectionIndexUpdate()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "CollectionIndexUpdate: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Updating Collection Index..."
    Application.ScreenUpdating = False

    UpdateCollectionIndex

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "CollectionIndexUpdate: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Collection index updated.", vbInformation, "Complete"
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' UpdateCollectionIndex
' Collect header info from DOC- sheets and write to
' UI_CollectionIndex Tbl:CollectionIndex
' ============================================
Private Sub UpdateCollectionIndex()
    If Not SheetExists(SHEET_UI_COLLECTION_INDEX) Then
        LogWarn TOOL_NAME, "Sheet not found: " & SHEET_UI_COLLECTION_INDEX
        Exit Sub
    End If

    Dim wsIdx As Worksheet
    Set wsIdx = ThisWorkbook.Worksheets(SHEET_UI_COLLECTION_INDEX)

    Dim markerRow As Long
    markerRow = FindTblStartRow(wsIdx, TBL_COLLECTION_INDEX)
    If markerRow = 0 Then
        LogWarn TOOL_NAME, TBL_MARKER_PREFIX & TBL_COLLECTION_INDEX & " not found"
        Exit Sub
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim headers As Variant
    headers = GetTableHeaders(wsIdx, headerRow)

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ' Collect from DOC- sheets
    Dim collections As Collection
    Set collections = CollectAllCollections(headers)

    ' Sort by collection_id
    Dim sortedData As Collection
    Set sortedData = SortByField(collections, "collection_id")

    ' Clear + write
    ClearTableData wsIdx, headerRow, colCount

    Dim dataRow As Long
    dataRow = headerRow + 1

    Dim rowNum As Long
    rowNum = 0

    Dim item As Object
    For Each item In sortedData
        rowNum = rowNum + 1
        item(COL_NO) = rowNum
        WriteTableRow wsIdx, dataRow, headers, item, COL_SHEET_NAME
        dataRow = dataRow + 1
    Next item

    LogInfo TOOL_NAME, "CollectionIndex: " & rowNum & " collections"
End Sub

' ============================================
' CollectAllSheets
' Collect info from all sheets for UI_SheetIndex
' ============================================
Private Function CollectAllSheets(targetColumns As Variant) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Skip dashboard itself
        If ws.Name = SHEET_UI_DASHBOARD Then GoTo NextSheet

        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        dict(COL_SHEET_NAME) = ws.Name

        ' Determine role from prefix
        Dim role As String
        role = GetRoleFromPrefix(ws.Name)
        If HasColumn(targetColumns, "role") Then
            dict("role") = role
        End If

        ' Try to read IndexHeader for note
        Dim markerRow As Long
        markerRow = FindTblStartRow(ws, TBL_INDEX_HEADER)
        If markerRow > 0 Then
            Dim headerInfo As Object
            Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)
            If headerInfo.Exists("note") And HasColumn(targetColumns, "note") Then
                dict("note") = headerInfo("note")
            End If
        End If

        result.Add dict

NextSheet:
    Next ws

    Set CollectAllSheets = result
End Function

' ============================================
' CollectAllCollections
' Collect header info from DOC- sheets
' ============================================
Private Function CollectAllCollections(targetColumns As Variant) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim docSheets As Collection
    Set docSheets = FilterSheetsByPrefix(PREFIX_COLLECTION)

    Dim sheetName As Variant
    For Each sheetName In docSheets
        ' Skip template
        If Left(CStr(sheetName), Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
            GoTo NextColl
        End If

        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))

        ' Find DOC_HeaderInfo
        Dim markerRow As Long
        markerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
        If markerRow = 0 Then
            LogWarn TOOL_NAME, TBL_MARKER_PREFIX & TBL_DOC_HEADER_INFO & " not found in " & sheetName
            GoTo NextColl
        End If

        Dim headerInfo As Object
        Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        dict(COL_SHEET_NAME) = CStr(sheetName)

        ' Map HeaderInfo keys to target columns
        ' HeaderInfo uses collection_ prefix; CollectionIndex columns may not.
        ' Try both the target column name and collection_ prefixed version.
        Dim i As Long
        Dim colName As String
        For i = LBound(targetColumns) To UBound(targetColumns)
            colName = targetColumns(i)
            If colName = COL_NO Or colName = COL_SHEET_NAME Then GoTo NextCol

            ' Special: doc_count -- count documents in DOC_DocumentList
            If colName = "doc_count" Then
                dict("doc_count") = CountDocuments(ws)
                GoTo NextCol
            End If

            ' Map collection_id from sheet name
            If colName = "collection_id" Then
                dict("collection_id") = CStr(sheetName)
                GoTo NextCol
            End If

            ' Direct mapping from header info (try exact key first, then collection_ prefixed)
            If headerInfo.Exists(colName) Then
                dict(colName) = headerInfo(colName)
            ElseIf headerInfo.Exists("collection_" & colName) Then
                dict(colName) = headerInfo("collection_" & colName)
            End If

NextCol:
        Next i

        result.Add dict
        LogInfo TOOL_NAME, "Collected: " & sheetName

NextColl:
    Next sheetName

    Set CollectAllCollections = result
End Function

' ============================================
' CountDocuments
' Count non-empty rows in DOC_DocumentList
' ============================================
Private Function CountDocuments(ws As Worksheet) As Long
    CountDocuments = 0

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_DOCUMENT_LIST)
    If markerRow = 0 Then Exit Function

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim r As Long
    For r = headerRow + 1 To headerRow + 300
        If IsEmpty(ws.Cells(r, 1).Value) Then Exit For
        CountDocuments = CountDocuments + 1
    Next r
End Function

' ============================================
' SortByField
' Sort collection of Dictionary by specified field
' ============================================
Private Function SortByField(data As Collection, fieldName As String) As Collection
    Dim result As Collection
    Set result = New Collection

    If data.Count = 0 Then
        Set SortByField = result
        Exit Function
    End If

    ' Convert to array
    Dim arr() As Variant
    ReDim arr(1 To data.Count)

    Dim i As Long
    For i = 1 To data.Count
        Set arr(i) = data(i)
    Next i

    ' Bubble sort
    Dim j As Long
    Dim swapped As Boolean
    Dim temp As Object
    Dim val1 As String, val2 As String

    For i = 1 To data.Count - 1
        swapped = False
        For j = 1 To data.Count - i
            val1 = ""
            val2 = ""
            If arr(j).Exists(fieldName) Then val1 = CStr(arr(j)(fieldName))
            If arr(j + 1).Exists(fieldName) Then val2 = CStr(arr(j + 1)(fieldName))

            If val1 > val2 Then
                Set temp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = temp
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For
    Next i

    For i = 1 To data.Count
        result.Add arr(i)
    Next i

    Set SortByField = result
End Function

' ============================================
' GetRoleFromPrefix
' Derive role string from sheet name prefix
' ============================================
Private Function GetRoleFromPrefix(sheetName As String) As String
    If Left(sheetName, Len(PREFIX_COLLECTION)) = PREFIX_COLLECTION Then
        GetRoleFromPrefix = "collection"
    ElseIf Left(sheetName, Len(PREFIX_UI)) = PREFIX_UI Then
        GetRoleFromPrefix = "ui"
    ElseIf Left(sheetName, Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
        GetRoleFromPrefix = "template"
    ElseIf Left(sheetName, Len(PREFIX_DEFINITION)) = PREFIX_DEFINITION Then
        GetRoleFromPrefix = "definition"
    ElseIf Left(sheetName, Len(PREFIX_LOG)) = PREFIX_LOG Then
        GetRoleFromPrefix = "log"
    ElseIf Left(sheetName, Len(PREFIX_INDEX)) = PREFIX_INDEX Then
        GetRoleFromPrefix = "index"
    ElseIf Left(sheetName, Len(PREFIX_MASTER)) = PREFIX_MASTER Then
        GetRoleFromPrefix = "master"
    Else
        GetRoleFromPrefix = ""
    End If
End Function

' ============================================
' HasColumn
' Check if column name exists in headers array
' ============================================
Private Function HasColumn(headers As Variant, colName As String) As Boolean
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If headers(i) = colName Then
            HasColumn = True
            Exit Function
        End If
    Next i
    HasColumn = False
End Function
