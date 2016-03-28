VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal l As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hDb                   As Long
Private m_lLastErr              As Long

'=========================================================================
' Properties
'=========================================================================

Property Get LastErr() As Long
    LastErr = m_lLastErr
End Property

Private Property Let pvResult(ByVal lValue As Long)
    m_lLastErr = lValue
    If lValue <> SQLITE_OK Then
        Err.Raise vbObjectError, , pvToString(vbsqlite3_errmsg(m_hDb))
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function OpenDb(FileName As String, Optional ByVal CreateIfNotExists As Boolean) As Boolean
    Dim baName()        As Byte
    
    If m_hDb <> 0 Then
        Call vbsqlite3_close_v2(m_hDb)
        m_hDb = 0
    End If
    baName = pvToUtf8(FileName)
    pvResult = vbsqlite3_open_v2(VarPtr(baName(0)), VarPtr(m_hDb), SQLITE_OPEN_READWRITE Or IIf(CreateIfNotExists, SQLITE_OPEN_CREATE, 0), 0)
    '--- success
    OpenDb = True
End Function

Public Function CloseDb() As Boolean
    If m_hDb <> 0 Then
        pvResult = vbsqlite3_close_v2(m_hDb)
        '--- success
        CloseDb = True
    End If
End Function

Public Function Execute(Sql As String, Optional Error As String) As Boolean
    Dim baSql()         As Byte
    Dim lPtrError       As Long
    Dim lResult         As Long
    
    baSql = pvToUtf8(Sql)
    lResult = vbsqlite3_exec(m_hDb, VarPtr(baSql(0)), 0, 0, VarPtr(lPtrError))
    Error = pvToString(lPtrError)
    Call vbsqlite3_free(lPtrError)
    lPtrError = 0
    pvResult = lResult
    '--- success (or failure)
    Execute = LenB(Error) = 0
End Function

Public Function Query(Sql As String) As cRecordset
    
End Function

'= private ===============================================================

Private Function pvToUtf8(sText As String) As Byte()
    Dim baBuffer()      As Byte
    Dim lSize           As Long
    
    ReDim baBuffer(0 To 4 * Len(sText)) As Byte
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baBuffer(0), UBound(baBuffer), 0, 0)
    ReDim Preserve baBuffer(0 To lSize) As Byte
    pvToUtf8 = baBuffer
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlenA(lPtr), Chr$(0))
        Call CopyMemory(ByVal pvToString, ByVal lPtr, Len(pvToString))
    End If
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Call vbsqlite3_win32_set_directory(SQLITE_WIN32_TEMP_DIRECTORY_TYPE, StrPtr(Environ$("TEMP")))
End Sub

Private Sub Class_Terminate()
    If m_hDb <> 0 Then
        Call vbsqlite3_close_v2(m_hDb)
        m_hDb = 0
    End If
End Sub
