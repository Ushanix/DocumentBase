Option Explicit

' ============================================
' Module   : Pst_RefreshDocumentList
' Layer    : Presentation
' Purpose  : Recalculate derived values in DOC_DocumentList
'            - doc_type_prefix (from DEF_DocType)
'            - document_id (collection_id + prefix + sequential number)
'            - Auto-fill: role (default "docs"), created/updated (today)
' Version  : 1.2.0
' Created  : 2026-03-22
' Updated  : 2026-03-24 — Auto-fill role/created/updated; collection_ prefix
' ============================================

Private Const TOOL_NAME As String = "RefreshDocumentList"

' Column names
Private Const COL_NO As String = "no"
Private Const COL_DOC_TYPE As String = "doc_type"
Private Const COL_DOC_TYPE_PREFIX As String = "doc_type_prefix"
Private Const COL_DOCUMENT_ID As String = "document_id"
Private Const COL_ROLE As String = "role"
Private Const COL_CREATED As String = "created"
Private Const COL_UPDATED As String = "updated"

' Default values for auto-fill
Private Const DEFAULT_ROLE As String = "docs"

' ============================================
' RefreshDocumentList
' Refresh derived values for the active DOC- sheet (single)
' ============================================
Public Sub RefreshDocumentList()
    On Error GoTo EH

    Dim currentSheet As String
    currentSheet = ActiveSheet.Name

    If Left(currentSheet, Len(PREFIX_COLLECTION)) <> PREFIX_COLLECTION Then
        MsgBox "Please run from a DOC-* sheet.", vbExclamation, "Error"
        Exit Sub
    End If

    If Left(currentSheet, Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
        MsgBox "Cannot refresh template sheet.", vbExclamation, "Error"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing document list..."

    Dim updatedRows As Long
    updatedRows = RefreshSheet(currentSheet)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Refresh completed." & vbCrLf & _
           updatedRows & " rows updated.", vbInformation, "Complete"
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' RefreshAll
' Refresh derived values for ALL DOC- sheets (batch)
' Called from UI_Dashboard action: refresh_all
' ============================================
Public Sub RefreshAll()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "RefreshAll: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Refreshing all collections..."
    Application.ScreenUpdating = False

    Dim docSheets As Collection
    Set docSheets = FilterSheetsByPrefix(PREFIX_COLLECTION)

    Dim successCount As Long
    Dim totalRows As Long
    successCount = 0
    totalRows = 0

    Dim sheetName As Variant
    For Each sheetName In docSheets
        ' Skip template
        If Left(CStr(sheetName), Len(PREFIX_TEMPLATE)) = PREFIX_TEMPLATE Then
            GoTo NextSheet
        End If

        Application.StatusBar = "Refreshing: " & CStr(sheetName) & "..."

        Dim rows As Long
        rows = RefreshSheet(CStr(sheetName))

        If rows >= 0 Then
            successCount = successCount + 1
            totalRows = totalRows + rows
        End If

NextSheet:
    Next sheetName

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "RefreshAll: Completed (" & successCount & " sheets, " & totalRows & " rows)"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Refresh all completed." & vbCrLf & _
           successCount & " collections, " & totalRows & " rows updated.", _
           vbInformation, "Complete"
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "RefreshAll Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' RefreshSheet (Private)
' Core logic: refresh derived values for a single sheet
'
' Args:
'   sheetName: Name of the DOC- sheet
'
' Returns:
'   Number of rows updated, or -1 on error
' ============================================
Private Function RefreshSheet(sheetName As String) As Long
    On Error GoTo EH

    RefreshSheet = -1

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    ' Get collection_id from HeaderInfo
    Dim collectionId As String
    collectionId = GetCollectionId(ws)
    If Len(collectionId) = 0 Then
        collectionId = sheetName
    End If

    LogInfo TOOL_NAME, "Refreshing: " & sheetName & " (id=" & collectionId & ")"

    ' Load doc_type -> id_prefix mapping
    Dim prefixMap As Object
    Set prefixMap = LoadDocTypePrefixMap()

    ' Find DOC_DocumentList
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_DOCUMENT_LIST)
    If markerRow = 0 Then
        LogWarn TOOL_NAME, "DOC_DocumentList not found in " & sheetName
        RefreshSheet = 0
        Exit Function
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim headers As Variant
    headers = GetTableHeaders(ws, headerRow)

    Dim colDocType As Long
    colDocType = GetColumnIndex(headers, COL_DOC_TYPE)

    Dim colPrefix As Long
    colPrefix = GetColumnIndex(headers, COL_DOC_TYPE_PREFIX)

    Dim colDocId As Long
    colDocId = GetColumnIndex(headers, COL_DOCUMENT_ID)

    Dim colNo As Long
    colNo = GetColumnIndex(headers, COL_NO)

    Dim colTitle As Long
    colTitle = GetColumnIndex(headers, "title")

    Dim colRole As Long
    colRole = GetColumnIndex(headers, COL_ROLE)

    Dim colCreated As Long
    colCreated = GetColumnIndex(headers, COL_CREATED)

    Dim colUpdated As Long
    colUpdated = GetColumnIndex(headers, COL_UPDATED)

    If colDocType = 0 Or colDocId = 0 Or colTitle = 0 Then
        LogWarn TOOL_NAME, "Required columns not found in " & sheetName
        RefreshSheet = 0
        Exit Function
    End If

    Dim today As String
    today = Format(Date, "yyyy-mm-dd")

    ' Scan rows: title が入力されている行を有効行とする
    ' no, doc_type_prefix, document_id を自動生成
    ' role が空なら "docs" を自動入力
    ' created/updated が空なら今日を自動入力
    Dim typeCounters As Object
    Set typeCounters = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim updatedRows As Long
    updatedRows = 0

    Dim rowSeq As Long
    rowSeq = 0

    For r = headerRow + 1 To headerRow + 300
        ' title の有無で行の存在を判定
        Dim titleVal As Variant
        titleVal = ws.Cells(r, colTitle).Value

        If IsEmpty(titleVal) Or Len(Trim(CStr(titleVal))) = 0 Then
            ' title が空の行 — no と導出値もクリア
            If colNo > 0 And Not IsEmpty(ws.Cells(r, colNo).Value) Then
                ws.Cells(r, colNo).Value = Empty
            End If
            If Not IsEmpty(ws.Cells(r, colDocId).Value) Then
                ws.Cells(r, colDocId).Value = Empty
            End If
            If colPrefix > 0 And Not IsEmpty(ws.Cells(r, colPrefix).Value) Then
                ws.Cells(r, colPrefix).Value = Empty
            End If
            ' 連続空行で終了判定（2行連続空なら終了）
            Dim nextTitle As Variant
            If r + 1 <= headerRow + 300 Then
                nextTitle = ws.Cells(r + 1, colTitle).Value
            Else
                nextTitle = Empty
            End If
            If IsEmpty(nextTitle) Or Len(Trim(CStr(nextTitle))) = 0 Then
                Exit For
            End If
            GoTo NextRow
        End If

        ' --- title がある行: no を自動採番 ---
        rowSeq = rowSeq + 1
        If colNo > 0 Then
            ws.Cells(r, colNo).Value = rowSeq
        End If

        ' --- role が空なら "docs" を自動入力 ---
        If colRole > 0 Then
            Dim roleVal As Variant
            roleVal = ws.Cells(r, colRole).Value
            If IsEmpty(roleVal) Or Len(Trim(CStr(roleVal))) = 0 Then
                ws.Cells(r, colRole).Value = DEFAULT_ROLE
            End If
        End If

        ' --- created が空なら今日を自動入力 ---
        If colCreated > 0 Then
            Dim createdVal As Variant
            createdVal = ws.Cells(r, colCreated).Value
            If IsEmpty(createdVal) Or Len(Trim(CStr(createdVal))) = 0 Then
                ws.Cells(r, colCreated).Value = today
            End If
        End If

        ' --- updated が空なら今日を自動入力 ---
        If colUpdated > 0 Then
            Dim updatedVal As Variant
            updatedVal = ws.Cells(r, colUpdated).Value
            If IsEmpty(updatedVal) Or Len(Trim(CStr(updatedVal))) = 0 Then
                ws.Cells(r, colUpdated).Value = today
            End If
        End If

        ' --- doc_type から prefix を検索 ---
        Dim docType As String
        docType = ""
        Dim docTypeVal As Variant
        docTypeVal = ws.Cells(r, colDocType).Value
        If Not IsEmpty(docTypeVal) Then
            docType = Trim(CStr(docTypeVal))
        End If

        If Len(docType) = 0 Then
            ' doc_type 未入力でも no は付与済み、prefix/id は空にする
            If colPrefix > 0 Then ws.Cells(r, colPrefix).Value = Empty
            ws.Cells(r, colDocId).Value = Empty
            updatedRows = updatedRows + 1
            GoTo NextRow
        End If

        ' Look up prefix
        Dim prefix As String
        prefix = ""
        If prefixMap.Exists(docType) Then
            prefix = CStr(prefixMap(docType))
        End If

        ' Update doc_type_prefix
        If colPrefix > 0 Then
            ws.Cells(r, colPrefix).Value = prefix
        End If

        ' Count for sequential numbering
        If Not typeCounters.Exists(docType) Then
            typeCounters(docType) = 0
        End If
        typeCounters(docType) = typeCounters(docType) + 1

        ' Generate document_id
        Dim seqNum As Long
        seqNum = typeCounters(docType)

        Dim docId As String
        If Len(prefix) > 0 Then
            docId = collectionId & "-" & prefix & Format(seqNum, "00")
        Else
            docId = collectionId & "-" & Format(seqNum, "00")
        End If

        ws.Cells(r, colDocId).Value = docId
        updatedRows = updatedRows + 1

NextRow:
    Next r

    ' Update HeaderInfo.collection_updated
    Dim headerMarkerRow As Long
    headerMarkerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
    If headerMarkerRow > 0 Then
        UpdateKeyValueTable ws, headerMarkerRow + 1, "collection_updated", Format(Date, "yyyy-mm-dd")
    End If

    LogInfo TOOL_NAME, "  " & sheetName & ": " & updatedRows & " rows"
    RefreshSheet = updatedRows
    Exit Function

EH:
    LogError TOOL_NAME, "Error refreshing " & sheetName & ": " & Err.Description
    RefreshSheet = -1
End Function

' ============================================
' GetCollectionId
' ============================================
Private Function GetCollectionId(ws As Worksheet) As String
    GetCollectionId = ""

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
    If markerRow = 0 Then Exit Function

    Dim headerInfo As Object
    Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

    ' collection_id key is not prefixed with collection_ as it serves as the primary identifier
    If headerInfo.Exists("collection_id") Then
        GetCollectionId = CStr(headerInfo("collection_id"))
    End If
End Function

' ============================================
' LoadDocTypePrefixMap
' ============================================
Private Function LoadDocTypePrefixMap() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim sheetName As String
    sheetName = "DEF_DocType"

    If Not SheetExists(sheetName) Then
        Set LoadDocTypePrefixMap = dict
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, "DEF_DocTypeData")
    If markerRow = 0 Then
        Set LoadDocTypePrefixMap = dict
        Exit Function
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim headers As Variant
    headers = GetTableHeaders(ws, headerRow)

    Dim colValue As Long
    colValue = GetColumnIndex(headers, "value")

    Dim colIdPrefix As Long
    colIdPrefix = GetColumnIndex(headers, "id_prefix")

    If colValue = 0 Or colIdPrefix = 0 Then
        Set LoadDocTypePrefixMap = dict
        Exit Function
    End If

    Dim r As Long
    For r = headerRow + 1 To headerRow + 100
        Dim val As Variant
        val = ws.Cells(r, colValue).Value
        If IsEmpty(val) Then Exit For

        Dim prefix As Variant
        prefix = ws.Cells(r, colIdPrefix).Value

        If Not IsEmpty(prefix) Then
            dict(CStr(val)) = CStr(prefix)
        End If
    Next r

    Set LoadDocTypePrefixMap = dict
End Function
