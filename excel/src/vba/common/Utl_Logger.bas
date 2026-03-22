Option Explicit

' ============================================
' Module   : Utl_Logger
' Layer    : Common / Utility
' Purpose  : Logging to Immediate Window
' Version  : 1.0.0
' Created  : 2026-03-22
' Note     : Simplified from FlowBase Utl_Logger
' ============================================

Public Sub LogInfo(ByVal toolName As String, ByVal message As String)
    Debug.Print "[INFO]  " & toolName & " | " & message
End Sub

Public Sub LogDebug(ByVal toolName As String, ByVal message As String)
    Debug.Print "[DEBUG] " & toolName & " | " & message
End Sub

Public Sub LogWarn(ByVal toolName As String, ByVal message As String)
    Debug.Print "[WARN]  " & toolName & " | " & message
End Sub

Public Sub LogError(ByVal toolName As String, ByVal message As String)
    Debug.Print "[ERROR] " & toolName & " | " & message
End Sub
