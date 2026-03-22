Option Explicit

' ============================================
' Module   : Utl_Sheet
' Layer    : Common / Utility
' Purpose  : Sheet operations (filter, sort, copy)
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Ported from FlowBase Utl_Sheet
' ============================================

' ============================================
' FilterSheetsByPrefix
' Get all sheet names starting with specified prefix
'
' Args:
'   prefix: Prefix to filter (e.g., "DOC-")
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   Collection of sheet names
' ============================================
Public Function FilterSheetsByPrefix(prefix As String, _
                                      Optional wb As Workbook = Nothing) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    Dim ws As Worksheet
    For Each ws In targetWb.Worksheets
        If Left(ws.Name, Len(prefix)) = prefix Then
            result.Add ws.Name
        End If
    Next ws

    Set FilterSheetsByPrefix = result
End Function

' ============================================
' SheetExists
' Check if sheet exists in workbook
' ============================================
Public Function SheetExists(sheetName As String, _
                             Optional wb As Workbook = Nothing) As Boolean
    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' ============================================
' CountSheetsByPrefix
' Count sheets starting with specified prefix
' ============================================
Public Function CountSheetsByPrefix(prefix As String, _
                                     Optional wb As Workbook = Nothing) As Long
    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    Dim cnt As Long
    cnt = 0

    Dim ws As Worksheet
    For Each ws In targetWb.Worksheets
        If Left(ws.Name, Len(prefix)) = prefix Then
            cnt = cnt + 1
        End If
    Next ws

    CountSheetsByPrefix = cnt
End Function

' ============================================
' CopySheet
' Copy template sheet to new sheet
'
' Args:
'   templateName: Template sheet name
'   newName: New sheet name
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   New worksheet, or Nothing if failed
' ============================================
Public Function CopySheet(templateName As String, _
                           newName As String, _
                           Optional wb As Workbook = Nothing) As Worksheet
    On Error GoTo ErrHandler

    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    If Not SheetExists(templateName, targetWb) Then
        Set CopySheet = Nothing
        Exit Function
    End If

    If SheetExists(newName, targetWb) Then
        Set CopySheet = Nothing
        Exit Function
    End If

    Dim templateWs As Worksheet
    Set templateWs = targetWb.Worksheets(templateName)

    templateWs.Copy After:=targetWb.Worksheets(targetWb.Worksheets.Count)

    Dim newWs As Worksheet
    Set newWs = targetWb.Worksheets(targetWb.Worksheets.Count)

    RenameTablesInSheet newWs
    newWs.Name = newName

    Set CopySheet = newWs
    Exit Function

ErrHandler:
    Set CopySheet = Nothing
End Function

' ============================================
' RenameTablesInSheet
' Rename all ListObjects in sheet to unique names
' ============================================
Public Sub RenameTablesInSheet(ws As Worksheet)
    On Error Resume Next

    Dim lo As ListObject
    Dim oldName As String
    Dim newName As String
    Dim suffix As Long

    For Each lo In ws.ListObjects
        oldName = lo.Name
        suffix = 1

        Do
            newName = oldName & "_" & suffix
            Err.Clear
            lo.Name = newName
            If Err.Number = 0 Then Exit Do
            suffix = suffix + 1
            If suffix > 100 Then Exit Do
        Loop
    Next lo

    On Error GoTo 0
End Sub

' ============================================
' LoadPrefixSortOrder
' Load prefix sort order from DEF_SheetPrefix
'
' Returns:
'   Dictionary with {prefix: sort_order}
' ============================================
Public Function LoadPrefixSortOrder(Optional wb As Workbook = Nothing) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    If Not SheetExists(SHEET_DEF_SHEET_PREFIX, targetWb) Then
        Set LoadPrefixSortOrder = dict
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = targetWb.Worksheets(SHEET_DEF_SHEET_PREFIX)

    Dim headerRow As Long
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_SHEET_PREFIX)
    If markerRow > 0 Then
        headerRow = markerRow + 1
    Else
        headerRow = 1
    End If

    Dim prefixCol As Long, orderCol As Long
    Dim col As Long
    Dim headerVal As Variant

    prefixCol = 0
    orderCol = 0

    For col = 1 To 20
        headerVal = ws.Cells(headerRow, col).Value
        If Not IsEmpty(headerVal) Then
            Select Case CStr(headerVal)
                Case "sheet_prefix": prefixCol = col
                Case "sort_order": orderCol = col
            End Select
        End If
    Next col

    If prefixCol = 0 Or orderCol = 0 Then
        Set LoadPrefixSortOrder = dict
        Exit Function
    End If

    Dim row As Long
    Dim prefix As Variant
    Dim order As Variant
    Dim orderVal As Long

    For row = headerRow + 1 To headerRow + 100
        prefix = ws.Cells(row, prefixCol).Value
        If IsEmpty(prefix) Or Trim(CStr(prefix)) = "" Then Exit For

        order = ws.Cells(row, orderCol).Value
        On Error Resume Next
        orderVal = CLng(order)
        If Err.Number <> 0 Then orderVal = DEFAULT_SORT_ORDER
        On Error GoTo 0

        dict(CStr(prefix)) = orderVal
    Next row

    Set LoadPrefixSortOrder = dict
End Function
