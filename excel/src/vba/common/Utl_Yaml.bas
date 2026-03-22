Option Explicit

' ============================================
' Module   : Utl_Yaml
' Layer    : Common / Utility
' Purpose  : Minimal YAML serializer/parser for DataIO
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Simplified subset — supports flat key-value and
'            tabular data only (no nested structures)
'
' YAML format:
'   meta:
'     key: value
'   tables:
'     TableName:
'       type: key_value | tabular
'       data:            (for key_value)
'         key: value
'       headers:         (for tabular)
'         - col1
'         - col2
'       rows:            (for tabular)
'         - col1: val1
'           col2: val2
' ============================================

' ============================================
' SerializeToYaml
' Convert meta + tables dictionaries to YAML string
' ============================================
Public Function SerializeToYaml(meta As Object, tablesDict As Object) As String
    Dim lines As Collection
    Set lines = New Collection

    ' --- meta section ---
    lines.Add "meta:"
    Dim mKey As Variant
    For Each mKey In meta.Keys
        lines.Add "  " & CStr(mKey) & ": " & YamlEscape(CStr(meta(mKey)))
    Next mKey

    ' --- tables section ---
    lines.Add "tables:"

    Dim tblName As Variant
    For Each tblName In tablesDict.Keys
        Dim tblInfo As Object
        Set tblInfo = tablesDict(tblName)

        Dim tblType As String
        tblType = CStr(tblInfo("type"))

        lines.Add "  " & CStr(tblName) & ":"
        lines.Add "    type: " & tblType

        If tblType = "key_value" Then
            lines.Add "    data:"
            Dim kvData As Object
            Set kvData = tblInfo("data")
            Dim kvKey As Variant
            For Each kvKey In kvData.Keys
                lines.Add "      " & CStr(kvKey) & ": " & YamlEscape(FormatVariant(kvData(kvKey)))
            Next kvKey

        ElseIf tblType = "tabular" Then
            Dim headers As Variant
            headers = tblInfo("headers")

            lines.Add "    headers:"
            Dim h As Long
            For h = LBound(headers) To UBound(headers)
                lines.Add "      - " & headers(h)
            Next h

            lines.Add "    rows:"
            Dim rows As Collection
            Set rows = tblInfo("rows")

            Dim row As Object
            For Each row In rows
                Dim firstCol As Boolean
                firstCol = True
                For h = LBound(headers) To UBound(headers)
                    Dim hdr As String
                    hdr = headers(h)
                    Dim val As String
                    val = ""
                    If row.Exists(hdr) Then val = FormatVariant(row(hdr))

                    If firstCol Then
                        lines.Add "      - " & hdr & ": " & YamlEscape(val)
                        firstCol = False
                    Else
                        lines.Add "        " & hdr & ": " & YamlEscape(val)
                    End If
                Next h
            Next row
        End If
    Next tblName

    ' Join
    Dim result As String
    Dim line As Variant
    For Each line In lines
        result = result & line & vbLf
    Next line

    SerializeToYaml = result
End Function

' ============================================
' ParseYamlFile
' Parse YAML file into Dictionary structure:
'   { "meta" -> Dictionary, "tables" -> Dictionary }
' ============================================
Public Function ParseYamlFile(filePath As String) As Object
    On Error GoTo EH

    Dim content As String
    content = ReadTextFile(filePath)

    If Len(content) = 0 Then
        Set ParseYamlFile = Nothing
        Exit Function
    End If

    Set ParseYamlFile = ParseYamlContent(content)
    Exit Function

EH:
    Set ParseYamlFile = Nothing
End Function

' ============================================
' ParseYamlMeta
' Parse only the meta section from a YAML file (fast)
' ============================================
Public Function ParseYamlMeta(filePath As String) As Object
    On Error GoTo EH

    Dim content As String
    content = ReadTextFile(filePath)

    If Len(content) = 0 Then
        Set ParseYamlMeta = Nothing
        Exit Function
    End If

    Dim lines() As String
    lines = Split(content, vbLf)

    Dim meta As Object
    Set meta = CreateObject("Scripting.Dictionary")

    Dim inMeta As Boolean
    inMeta = False

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String
        line = lines(i)

        ' Remove CR
        If Right(line, 1) = vbCr Then line = Left(line, Len(line) - 1)

        If Trim(line) = "meta:" Then
            inMeta = True
            GoTo NextMetaLine
        End If

        If inMeta Then
            ' End of meta: line with no leading space
            If Len(line) > 0 And Left(line, 1) <> " " Then
                Exit For
            End If

            ' Parse "  key: value"
            Dim trimmed As String
            trimmed = Trim(line)
            If Len(trimmed) = 0 Then GoTo NextMetaLine

            Dim colonPos As Long
            colonPos = InStr(trimmed, ": ")
            If colonPos > 0 Then
                Dim k As String
                k = Left(trimmed, colonPos - 1)
                Dim v As String
                v = Mid(trimmed, colonPos + 2)
                meta(k) = YamlUnescape(v)
            End If
        End If

NextMetaLine:
    Next i

    Set ParseYamlMeta = meta
    Exit Function

EH:
    Set ParseYamlMeta = Nothing
End Function

' ============================================
' ParseYamlContent
' Full parse of YAML content
' ============================================
Private Function ParseYamlContent(content As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim meta As Object
    Set meta = CreateObject("Scripting.Dictionary")

    Dim tablesDict As Object
    Set tablesDict = CreateObject("Scripting.Dictionary")

    Dim lines() As String
    lines = Split(content, vbLf)

    ' State machine
    Dim section As String    ' "meta", "tables"
    Dim tableName As String
    Dim tableType As String
    Dim subSection As String ' "data", "headers", "rows"
    Dim currentHeaders As Collection
    Dim currentRows As Collection
    Dim kvData As Object
    Dim currentRow As Object
    Dim isFirstRowCol As Boolean

    section = ""
    tableName = ""
    tableType = ""
    subSection = ""
    Set currentHeaders = New Collection
    Set currentRows = New Collection

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String
        line = lines(i)
        If Right(line, 1) = vbCr Then line = Left(line, Len(line) - 1)

        Dim trimmed As String
        trimmed = Trim(line)
        If Len(trimmed) = 0 Then GoTo NextParseLine

        ' Count leading spaces
        Dim indent As Long
        indent = Len(line) - Len(LTrim(line))

        ' Level 0: top sections
        If indent = 0 Then
            ' Flush previous table
            If Len(tableName) > 0 Then
                FlushTable tablesDict, tableName, tableType, kvData, currentHeaders, currentRows
                tableName = ""
            End If

            If trimmed = "meta:" Then
                section = "meta"
            ElseIf trimmed = "tables:" Then
                section = "tables"
            End If
            GoTo NextParseLine
        End If

        ' Level 1 (2 spaces): meta keys or table names
        If indent = 2 And section = "meta" Then
            Dim colonPos As Long
            colonPos = InStr(trimmed, ": ")
            If colonPos > 0 Then
                meta(Left(trimmed, colonPos - 1)) = YamlUnescape(Mid(trimmed, colonPos + 2))
            End If
            GoTo NextParseLine
        End If

        If indent = 2 And section = "tables" Then
            ' New table name (ends with :)
            If Right(trimmed, 1) = ":" Then
                ' Flush previous
                If Len(tableName) > 0 Then
                    FlushTable tablesDict, tableName, tableType, kvData, currentHeaders, currentRows
                End If
                tableName = Left(trimmed, Len(trimmed) - 1)
                tableType = ""
                subSection = ""
                Set kvData = CreateObject("Scripting.Dictionary")
                Set currentHeaders = New Collection
                Set currentRows = New Collection
                Set currentRow = Nothing
            End If
            GoTo NextParseLine
        End If

        ' Level 2 (4 spaces): table properties
        If indent = 4 And section = "tables" Then
            colonPos = InStr(trimmed, ": ")
            If colonPos > 0 Then
                Dim propKey As String
                propKey = Left(trimmed, colonPos - 1)
                Dim propVal As String
                propVal = Mid(trimmed, colonPos + 2)

                If propKey = "type" Then
                    tableType = propVal
                End If
            End If

            ' Sub-section markers
            If trimmed = "data:" Then subSection = "data"
            If trimmed = "headers:" Then subSection = "headers"
            If trimmed = "rows:" Then subSection = "rows"
            GoTo NextParseLine
        End If

        ' Level 3 (6 spaces): data content
        If indent = 6 Then
            If subSection = "data" Then
                colonPos = InStr(trimmed, ": ")
                If colonPos > 0 Then
                    kvData(Left(trimmed, colonPos - 1)) = YamlUnescape(Mid(trimmed, colonPos + 2))
                End If
            ElseIf subSection = "headers" Then
                If Left(trimmed, 2) = "- " Then
                    currentHeaders.Add Mid(trimmed, 3)
                End If
            ElseIf subSection = "rows" Then
                If Left(trimmed, 2) = "- " Then
                    ' New row
                    If Not currentRow Is Nothing Then
                        currentRows.Add currentRow
                    End If
                    Set currentRow = CreateObject("Scripting.Dictionary")
                    ' Parse first field
                    Dim rest As String
                    rest = Mid(trimmed, 3)
                    colonPos = InStr(rest, ": ")
                    If colonPos > 0 Then
                        currentRow(Left(rest, colonPos - 1)) = YamlUnescape(Mid(rest, colonPos + 2))
                    End If
                End If
            End If
            GoTo NextParseLine
        End If

        ' Level 4 (8 spaces): continuation of row fields
        If indent = 8 And subSection = "rows" Then
            If Not currentRow Is Nothing Then
                colonPos = InStr(trimmed, ": ")
                If colonPos > 0 Then
                    currentRow(Left(trimmed, colonPos - 1)) = YamlUnescape(Mid(trimmed, colonPos + 2))
                End If
            End If
        End If

NextParseLine:
    Next i

    ' Flush last row and table
    If Not currentRow Is Nothing Then
        currentRows.Add currentRow
    End If
    If Len(tableName) > 0 Then
        FlushTable tablesDict, tableName, tableType, kvData, currentHeaders, currentRows
    End If

    Set result("meta") = meta
    Set result("tables") = tablesDict

    Set ParseYamlContent = result
End Function

' ============================================
' FlushTable — save parsed table into tablesDict
' ============================================
Private Sub FlushTable(tablesDict As Object, tableName As String, tableType As String, _
                        kvData As Object, currentHeaders As Collection, currentRows As Collection)
    Dim tblInfo As Object
    Set tblInfo = CreateObject("Scripting.Dictionary")
    tblInfo("type") = tableType

    If tableType = "key_value" Then
        Set tblInfo("data") = kvData
    ElseIf tableType = "tabular" Then
        ' Convert headers Collection to Array
        Dim hdrs() As String
        If currentHeaders.Count > 0 Then
            ReDim hdrs(1 To currentHeaders.Count)
            Dim h As Long
            For h = 1 To currentHeaders.Count
                hdrs(h) = currentHeaders(h)
            Next h
        Else
            ReDim hdrs(1 To 1)
            hdrs(1) = ""
        End If
        tblInfo("headers") = hdrs
        Set tblInfo("rows") = currentRows
    End If

    Set tablesDict(tableName) = tblInfo
End Sub

' ============================================
' Helpers
' ============================================
Private Function YamlEscape(val As String) As String
    If Len(val) = 0 Then
        YamlEscape = "''"
        Exit Function
    End If
    If InStr(val, ":") > 0 Or InStr(val, "#") > 0 Or InStr(val, "'") > 0 Or Left(val, 1) = " " Then
        YamlEscape = """" & Replace(val, """", "\""") & """"
    Else
        YamlEscape = val
    End If
End Function

Private Function YamlUnescape(val As String) As String
    Dim s As String
    s = Trim(val)
    If Len(s) >= 2 Then
        If (Left(s, 1) = """" And Right(s, 1) = """") Or _
           (Left(s, 1) = "'" And Right(s, 1) = "'") Then
            s = Mid(s, 2, Len(s) - 2)
            s = Replace(s, "\""", """")
        End If
    End If
    If s = "''" Then s = ""
    YamlUnescape = s
End Function

Private Function FormatVariant(val As Variant) As String
    If IsEmpty(val) Or IsNull(val) Then
        FormatVariant = ""
    ElseIf IsDate(val) Then
        FormatVariant = Format(val, "yyyy-mm-dd")
    Else
        FormatVariant = CStr(val)
    End If
End Function
