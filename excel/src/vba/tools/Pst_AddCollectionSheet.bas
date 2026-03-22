Option Explicit

' ============================================
' Module   : Pst_AddCollectionSheet
' Layer    : Presentation
' Purpose  : Create new Collection sheet from template
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Adapted from FlowBase Pst_AddProjectSheet
' ============================================

Private Const TOOL_NAME As String = "AddCollectionSheet"

' ============================================
' AddCollectionSheet
' Create new DOC- sheet from TPL_DOC-CATEGORY-SEQ template
'
' Flow:
'   1. Read parameters from UI_AddSheet Tbl:AddCollection
'   2. Validate domain (required) and collection_name (required)
'   3. Get domain_code from DEF_CollectionDomain
'   4. Find max SEQ for that domain
'   5. Generate collection_id = DOC-<DOMAIN_CODE>-<SEQ>
'   6. Copy template, update DOC_HeaderInfo
'   7. Set auto values: status=active, created/updated=today
'   8. Clear AddCollection inputs
' ============================================
Public Sub AddCollectionSheet()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "AddCollectionSheet: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.ScreenUpdating = False

    ' --- 1. Read parameters ---
    If Not SheetExists(SHEET_UI_ADD_SHEET) Then
        LogError TOOL_NAME, "Sheet not found: " & SHEET_UI_ADD_SHEET
        MsgBox "Sheet not found: " & SHEET_UI_ADD_SHEET, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim wsAdd As Worksheet
    Set wsAdd = ThisWorkbook.Worksheets(SHEET_UI_ADD_SHEET)

    Dim params As Object
    Set params = ParseAddCollectionTable(wsAdd)

    ' --- 2. Validate required fields ---
    Dim collectionName As String
    collectionName = GetParamStr(params, "collection_name")
    If Len(collectionName) = 0 Then
        MsgBox "collection_name is required.", vbExclamation, "Validation Error"
        GoTo Cleanup
    End If

    ' --- 2b. Validate collection_name for folder safety ---
    Dim nameErrors As String
    nameErrors = ValidateCollectionName(collectionName)
    If Len(nameErrors) > 0 Then
        MsgBox "collection_name に問題があります。修正してください。" & vbCrLf & vbCrLf & _
               nameErrors & vbCrLf & vbCrLf & _
               "現在の値: " & collectionName, vbExclamation, "Collection Name Validation"
        GoTo Cleanup
    End If

    Dim domainValue As String
    domainValue = GetParamStr(params, "domain")
    If Len(domainValue) = 0 Then
        MsgBox "domain is required.", vbExclamation, "Validation Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "collection_name: " & collectionName
    LogInfo TOOL_NAME, "domain: " & domainValue

    ' --- 3. Get domain_code ---
    ' DEF_CollectionDomain uses value column directly as domain code
    ' Domain values like "Technology" are used; we need a short code for sheet name
    ' Use first N chars or lookup. For now, use the domain value as-is
    ' since DOC- sheet naming uses domain codes like TECH, ENV, MIND etc.
    Dim domainCode As String
    domainCode = GetDomainCode(domainValue)

    If Len(domainCode) = 0 Then
        LogError TOOL_NAME, "domain_code not resolved for: " & domainValue
        MsgBox "domain_code not resolved for: " & domainValue, vbExclamation, "Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "domain_code: " & domainCode

    ' --- 4. Find max SEQ ---
    Dim maxSeq As Long
    maxSeq = FindMaxSeq(domainCode)

    Dim newSeq As Long
    newSeq = maxSeq + 1

    LogInfo TOOL_NAME, "Max SEQ: " & maxSeq & " -> New SEQ: " & newSeq

    ' --- 5. Generate collection_id / sheet name ---
    Dim collectionId As String
    collectionId = PREFIX_COLLECTION & domainCode & "-" & Format(newSeq, "00")

    LogInfo TOOL_NAME, "collection_id: " & collectionId

    ' Validate sheet name
    Dim validationError As String
    validationError = ValidateSheetName(collectionId)
    If Len(validationError) > 0 Then
        MsgBox "Invalid sheet name: " & collectionId & vbCrLf & validationError, _
               vbExclamation, "Error"
        GoTo Cleanup
    End If

    If SheetExists(collectionId) Then
        MsgBox "Sheet already exists: " & collectionId, vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' --- 6. Copy template ---
    Application.StatusBar = "Creating collection sheet..."

    Dim templateName As String
    templateName = GetTemplateName()

    If Not SheetExists(templateName) Then
        MsgBox "Template not found: " & templateName, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim newWs As Worksheet
    Set newWs = CopySheet(templateName, collectionId)

    If newWs Is Nothing Then
        MsgBox "Failed to copy template.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "Template copied: " & collectionId

    ' --- 7. Update DOC_HeaderInfo ---
    Dim updated As Long
    updated = UpdateCollectionHeader(newWs, collectionId, params)
    LogInfo TOOL_NAME, "Updated " & updated & " fields in DOC_HeaderInfo"

    ' --- 8. Clear AddCollection inputs ---
    ClearAddCollectionValues wsAdd
    LogInfo TOOL_NAME, "Cleared AddCollection inputs"

    ' Activate new sheet
    newWs.Activate

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "AddCollectionSheet: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Collection sheet created." & vbCrLf & _
           "Sheet: " & collectionId, vbInformation, "Complete"
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
' ParseAddCollectionTable
' Read Tbl:AddCollection key-value table
' ============================================
Private Function ParseAddCollectionTable(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_ADD_COLLECTION)
    If markerRow = 0 Then
        LogError TOOL_NAME, TBL_MARKER_PREFIX & TBL_ADD_COLLECTION & " not found"
        Set ParseAddCollectionTable = dict
        Exit Function
    End If

    Set dict = ReadKeyValueTable(ws, markerRow + 1)
    LogInfo TOOL_NAME, "Parsed AddCollection: " & dict.Count & " parameters"

    Set ParseAddCollectionTable = dict
End Function

' ============================================
' GetDomainCode
' Resolve domain value to short code for sheet naming
'
' Looks up DEF_CollectionDomain for a matching value,
' then derives a short uppercase code.
'
' Strategy: Use first 4 chars uppercased as default,
' or lookup a code_column if available.
' ============================================
Private Function GetDomainCode(domainValue As String) As String
    ' Domain values are like "Technology", "Mind", "Self", "Env", etc.
    ' Sheet naming expects short codes: TECH, MIND, SELF, ENV, etc.
    '
    ' Convention: uppercase the value, truncate to reasonable length
    ' Special mappings for longer names
    Dim code As String

    Select Case domainValue
        Case "Technology"
            code = "TECH"
        Case "Relation"
            code = "REL"
        Case "Finance"
            code = "FIN"
        Case "Spirit"
            code = "SPRT"
        Case Else
            ' Use uppercase of value (works for Self, Mind, Art, Env, Career)
            code = UCase(domainValue)
            ' Cap at 5 characters
            If Len(code) > 5 Then
                code = Left(code, 5)
            End If
    End Select

    GetDomainCode = code
End Function

' ============================================
' FindMaxSeq
' Find maximum SEQ number for given domain code
'
' Scans for sheets named DOC-<domainCode>-<NN>
' ============================================
Private Function FindMaxSeq(domainCode As String) As Long
    Dim pattern As String
    pattern = PREFIX_COLLECTION & domainCode & "-"

    Dim maxSeq As Long
    maxSeq = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, Len(pattern)) = pattern Then
            Dim seqPart As String
            seqPart = Mid(ws.Name, Len(pattern) + 1)

            If IsNumeric(seqPart) Then
                Dim seq As Long
                seq = CLng(seqPart)
                If seq > maxSeq Then
                    maxSeq = seq
                End If
            End If
        End If
    Next ws

    LogInfo TOOL_NAME, "Max SEQ for " & pattern & "*: " & maxSeq
    FindMaxSeq = maxSeq
End Function

' ============================================
' GetTemplateName
' Get template sheet name from DEF_Parameter or use default
' ============================================
Private Function GetTemplateName() As String
    GetTemplateName = DEFAULT_COLLECTION_TEMPLATE

    If Not SheetExists(SHEET_DEF_PARAMETER) Then Exit Function

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)

    Dim result As Variant
    result = LookupTableValue(ws, TBL_DEF_PARAMETER, "name", "value", PARAM_COLLECTION_TEMPLATE)

    If Not IsEmpty(result) And Len(CStr(result)) > 0 Then
        GetTemplateName = CStr(result)
    End If
End Function

' ============================================
' UpdateCollectionHeader
' Update DOC_HeaderInfo in new collection sheet
'
' Sets:
'   - collection_id (auto)
'   - collection_name, summary, domain, related_project (from params)
'   - status = "active" (auto)
'   - created = today (auto)
'   - updated = today (auto)
'   - output_path = "" (empty, uses default)
' ============================================
Private Function UpdateCollectionHeader(ws As Worksheet, _
                                         collectionId As String, _
                                         params As Object) As Long
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_DOC_HEADER_INFO)
    If markerRow = 0 Then
        LogError TOOL_NAME, TBL_MARKER_PREFIX & TBL_DOC_HEADER_INFO & " not found"
        UpdateCollectionHeader = 0
        Exit Function
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim updated As Long
    updated = 0

    Dim today As String
    today = Format(Date, "yyyy-mm-dd")

    ' Auto values
    If UpdateKeyValueTable(ws, headerRow, "collection_id", collectionId) Then updated = updated + 1
    If UpdateKeyValueTable(ws, headerRow, "status", "active") Then updated = updated + 1
    If UpdateKeyValueTable(ws, headerRow, "created", today) Then updated = updated + 1
    If UpdateKeyValueTable(ws, headerRow, "updated", today) Then updated = updated + 1

    ' User input values
    Dim key As Variant
    For Each key In params.Keys
        Dim keyStr As String
        keyStr = CStr(key)

        Dim val As Variant
        val = params(key)

        ' Skip empty values
        If IsEmpty(val) Then GoTo NextParam
        If Len(Trim(CStr(val))) = 0 Then GoTo NextParam

        If UpdateKeyValueTable(ws, headerRow, keyStr, val) Then
            LogInfo TOOL_NAME, "  Set: " & keyStr & " = " & CStr(val)
            updated = updated + 1
        Else
            LogWarn TOOL_NAME, "  Key not found in HeaderInfo: " & keyStr
        End If

NextParam:
    Next key

    UpdateCollectionHeader = updated
End Function

' ============================================
' ClearAddCollectionValues
' Clear value column in AddCollection table
' ============================================
Private Sub ClearAddCollectionValues(ws As Worksheet)
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_ADD_COLLECTION)
    If markerRow = 0 Then Exit Sub

    ClearKeyValueTableValues ws, markerRow + 1
End Sub

' ============================================
' ClearKeyValueTableValues (local copy)
' Clear all values in key-value table (preserve keys)
' ============================================
Private Function ClearKeyValueTableValues(ws As Worksheet, _
                                           headerRow As Long, _
                                           Optional maxRows As Long = 100) As Long
    Dim i As Long
    Dim keyVal As Variant
    Dim cleared As Long
    cleared = 0

    For i = headerRow + 1 To headerRow + maxRows
        keyVal = ws.Cells(i, 1).Value
        If IsEmpty(keyVal) Or Trim(CStr(keyVal)) = "" Then Exit For

        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            ws.Cells(i, 2).Value = Empty
            cleared = cleared + 1
        End If
    Next i

    ClearKeyValueTableValues = cleared
End Function

' ============================================
' GetParamStr
' Safe string extraction from Dictionary
' ============================================
Private Function GetParamStr(dict As Object, key As String) As String
    GetParamStr = ""
    If dict.Exists(key) Then
        If Not IsEmpty(dict(key)) Then
            GetParamStr = Trim(CStr(dict(key)))
        End If
    End If
End Function

' ============================================
' ValidateSheetName
' Validate sheet name for Excel restrictions
' ============================================
Private Function ValidateSheetName(sheetName As String) As String
    Const MAX_LEN As Long = 31
    Const INVALID_CHARS As String = ":\/?*[]"

    ValidateSheetName = ""

    If Len(sheetName) > MAX_LEN Then
        ValidateSheetName = "Sheet name exceeds " & MAX_LEN & " characters."
        Exit Function
    End If

    If Len(sheetName) = 0 Then
        ValidateSheetName = "Sheet name cannot be empty."
        Exit Function
    End If

    Dim i As Long
    Dim c As String
    For i = 1 To Len(INVALID_CHARS)
        c = Mid(INVALID_CHARS, i, 1)
        If InStr(sheetName, c) > 0 Then
            ValidateSheetName = "Invalid character: " & c
            Exit Function
        End If
    Next i
End Function

' ============================================
' ValidateCollectionName
' Validate collection_name for folder-safe usage.
' Returns empty string if valid, or multi-line error description.
'
' Rules:
'   - Max 40 characters (folder name = collection_id + "_" + name)
'   - No filesystem-invalid characters: \ / : * ? " < > |
'   - No leading/trailing spaces (auto-trimmed, not an error)
'   - No control characters
'   - No consecutive spaces
'   - No period at end (Windows folder restriction)
' ============================================
Private Function ValidateCollectionName(collName As String) As String
    Const MAX_NAME_LEN As Long = 40
    Const INVALID_CHARS As String = "\/:*?""<>|"

    Dim errors As Collection
    Set errors = New Collection

    ' --- Length check ---
    If Len(collName) > MAX_NAME_LEN Then
        errors.Add "× " & MAX_NAME_LEN & "文字以内にしてください（現在 " & Len(collName) & " 文字）"
    End If

    ' --- Invalid characters ---
    Dim foundChars As String
    foundChars = ""
    Dim i As Long
    Dim c As String
    For i = 1 To Len(INVALID_CHARS)
        c = Mid(INVALID_CHARS, i, 1)
        If InStr(collName, c) > 0 Then
            foundChars = foundChars & " " & c
        End If
    Next i

    If Len(foundChars) > 0 Then
        errors.Add "× 使用できない文字が含まれています:" & foundChars
    End If

    ' --- Control characters (Unicode 0-31) ---
    ' AscW returns signed Integer; CJK characters (U+8000+) return negative.
    ' Only flag characters in the 0-31 range.
    For i = 1 To Len(collName)
        Dim charCode As Long
        charCode = AscW(Mid(collName, i, 1)) And &HFFFF&
        If charCode < 32 Then
            errors.Add "× 制御文字（改行・タブ等）は使用できません"
            Exit For
        End If
    Next i

    ' --- Consecutive spaces ---
    If InStr(collName, "  ") > 0 Then
        errors.Add "× 連続するスペースは使用できません"
    End If

    ' --- Trailing period ---
    If Right(collName, 1) = "." Then
        errors.Add "× 末尾にピリオドは使用できません（Windowsフォルダ制限）"
    End If

    ' --- Build result ---
    If errors.Count = 0 Then
        ValidateCollectionName = ""
    Else
        Dim result As String
        Dim errItem As Variant
        For Each errItem In errors
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & CStr(errItem)
        Next errItem
        ValidateCollectionName = result
    End If
End Function
