Attribute VB_Name = "Pst_ImportModules"
Option Explicit

' ============================================
' Module   : Pst_ImportModules
' Layer    : Presentation
' Purpose  : Import VBA modules from DocumentBase repository
'            Requires Ent_Import module from UniversalModelForVBA
' Version  : 1.0.0
' Created  : 2026-03-29
' ============================================

Private Const TOOL_NAME As String = "ImportModules"

' ============================================
' ImportDocumentBaseModules
' Import all VBA modules from DocumentBase repository.
' Folder structure:
'   excel/src/vba/common/   -> .bas (utilities, constants)
'   excel/src/vba/tools/    -> .bas (presentation layer)
' ============================================
Public Sub ImportDocumentBaseModules()
    On Error GoTo EH

    Dim basePath As String
    basePath = "C:\Dev\Github\DocumentBase\excel\src\vba"

    ' Confirm
    Dim msg As String
    msg = "DocumentBase VBA modules to import from:" & vbCrLf & _
          "  " & basePath & vbCrLf & vbCrLf & _
          "  common\  (bas)" & vbCrLf & _
          "  tools\   (bas)" & vbCrLf & vbCrLf & _
          "Existing modules will be overwritten." & vbCrLf & _
          "Sheet modules (ThisWorkbook, Sheet*) are skipped." & vbCrLf & vbCrLf & _
          "Continue?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Import DocumentBase Modules") <> vbYes Then
        Exit Sub
    End If

    ' Round 1: common (bas)
    Import_AllModulesEx _
        basDir:=basePath & "\common", _
        clsDir:="", _
        frmDir:="", _
        overwriteExisting:=True

    ' Round 2: tools (bas)
    Import_AllModulesEx _
        basDir:=basePath & "\tools", _
        clsDir:="", _
        frmDir:="", _
        overwriteExisting:=True

    Exit Sub

EH:
    MsgBox "Import error: " & Err.Description, vbCritical, "Import DocumentBase Modules"
End Sub
