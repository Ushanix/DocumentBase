Option Explicit

' ============================================
' Module   : Mod_Constants
' Layer    : Common
' Purpose  : Centralized constants for DocumentBase-VBA
' Version  : 1.0.0
' Created  : 2026-03-22
' ============================================

' ============================================
' Tbl Marker Prefix
' ============================================
Public Const TBL_MARKER_PREFIX As String = "Tbl:"

' ============================================
' Tbl Marker Names (PascalCase)
' Cell value = TBL_MARKER_PREFIX & MarkerName
' ============================================

' --- UI_Dashboard ---
Public Const TBL_INDEX_HEADER As String = "IndexHeader"
Public Const TBL_UI_OPERATIONS As String = "UI_Operations"
Public Const TBL_UI_STATUS As String = "UI_Status"
Public Const TBL_UI_SHEET_INDEX As String = "UI_SheetIndex"

' --- UI_CollectionIndex ---
Public Const TBL_ACTIONS As String = "Actions"
Public Const TBL_COLLECTION_INDEX As String = "CollectionIndex"

' --- UI_AddSheet ---
Public Const TBL_ADD_COLLECTION As String = "AddCollection"

' --- DEF_Parameter ---
Public Const TBL_DEF_PARAMETER As String = "DEF_Parameter"

' --- DEF_SheetPrefix ---
Public Const TBL_SHEET_PREFIX As String = "DEF_SheetPrefix"

' --- LOG_UpdateHistory ---
Public Const TBL_LOG_UPDATE_HISTORY As String = "LOG_UpdateHistory"

' --- DOC-* (Collection sheets) ---
Public Const TBL_DOC_HEADER_INFO As String = "DOC_HeaderInfo"
Public Const TBL_DOC_DOCUMENT_LIST As String = "DOC_DocumentList"
Public Const TBL_DOC_ACTION As String = "DOC_Action"

' --- TPL (Template) ---
Public Const TBL_TPL_HEADER_INFO As String = "TPL_HeaderInfo"

' ============================================
' Sheet Name Prefixes
' ============================================
Public Const PREFIX_COLLECTION As String = "DOC-"
Public Const PREFIX_TEMPLATE As String = "TPL_"
Public Const PREFIX_DEFINITION As String = "DEF_"
Public Const PREFIX_UI As String = "UI_"
Public Const PREFIX_LOG As String = "LOG_"
Public Const PREFIX_INDEX As String = "IDX_"
Public Const PREFIX_MASTER As String = "M_"

' ============================================
' Fixed Sheet Names
' ============================================
Public Const SHEET_UI_DASHBOARD As String = "UI_Dashboard"
Public Const SHEET_UI_COLLECTION_INDEX As String = "UI_CollectionIndex"
Public Const SHEET_UI_ADD_SHEET As String = "UI_AddSheet"
Public Const SHEET_DEF_PARAMETER As String = "DEF_Parameter"
Public Const SHEET_DEF_SHEET_PREFIX As String = "DEF_SheetPrefix"
Public Const SHEET_UI_DATA_IO As String = "UI_DataIO"

' ============================================
' DataIO Tbl Marker Names
' ============================================
Public Const TBL_DATA_IO_CONFIG As String = "DataIOConfig"
Public Const TBL_EXPORT_LIST As String = "ExportList"
Public Const TBL_IMPORT_LIST As String = "ImportList"
Public Const SHEET_LOG_UPDATE_HISTORY As String = "LOG_UpdateHistory"
Public Const SHEET_TPL_COLLECTION As String = "TPL_DOC-CATEGORY-SEQ"

' ============================================
' Parameter Keys (from DEF_Parameter)
' ============================================
Public Const PARAM_OUTPUT_ROOT As String = "OUTPUT_ROOT"
Public Const PARAM_OUTPUT_MODE As String = "OUTPUT_MODE"
Public Const PARAM_VERSION_DEFAULT As String = "VERSION_DEFAULT"
Public Const PARAM_COLLECTION_TEMPLATE As String = "COLLECTION_TEMPLATE"
Public Const PARAM_PREFIX_COLLECTION As String = "PREFIX_COLLECTION"
Public Const PARAM_DATA_EXPORT_PATH As String = "DATA_EXPORT_PATH"

' ============================================
' Default Values
' ============================================
Public Const DEFAULT_COLLECTION_TEMPLATE As String = "TPL_DOC-CATEGORY-SEQ"
Public Const DEFAULT_SORT_ORDER As Long = 9999

' ============================================
' Key-Value Table Names
' ============================================
Public Function GetKeyValueTables() As Variant
    GetKeyValueTables = Array( _
        "DOC_HeaderInfo", "IndexHeader", "AddCollection", "UI_Status")
End Function
