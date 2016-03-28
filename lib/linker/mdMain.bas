Attribute VB_Name = "mdMain"
Option Explicit

'=========================================================================
' API
'=========================================================================

Private Const INVALID_FILE_ATTRIBUTES       As Long = -1

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_COBJ_EXT              As String = "cobj"
Private Const STR_ORIGINAL_LINKER       As String = "vblink.exe"

'=========================================================================
' Functions
'=========================================================================

Public Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\", "\", vbNullString) & sFile
End Function

Public Function SplitArgs(sText As String) As Variant
    Dim oMatches        As Object
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = """([^""]*(?:""""[^""]*)*)""|([^ ]+)"
        Set oMatches = .Execute(sText)
        If oMatches.Count > 0 Then
            ReDim vRetVal(0 To oMatches.Count - 1) As String
            For lIdx = 0 To oMatches.Count - 1
                With oMatches(lIdx)
                    vRetVal(lIdx) = Replace$(.SubMatches(0) & .SubMatches(1), """""", """")
                End With
            Next
        Else
            vRetVal = Split(vbNullString)
        End If
    End With
    SplitArgs = vRetVal
End Function

Public Function ShellWait(sCommand As String, Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
    Call OutputDebugString("ShellWait: sCommand=" & sCommand & vbCrLf)
    ShellWait = CreateObject("WScript.Shell").Run(sCommand, WindowStyle, True)
End Function

Public Function ReadFile(sFileName As String) As String
    Dim nFile       As Integer
    Dim sContents   As String

    On Error GoTo QH
    '--- read contents
    nFile = FreeFile
    Open sFileName For Binary As #nFile
    sContents = Space(LOF(nFile))
    Get #nFile, , sContents
    Close #nFile
    '--- convert unicode text files
    If Left(sContents, 2) = "ÿþ" Then
        sContents = Mid(StrConv(sContents, vbFromUnicode), 2)
    End If
    ReadFile = sContents
QH:
End Function

Public Sub Main()
    Dim sCommand        As String
    Dim vElem           As Variant
    Dim lPos            As Long
    Dim sFile           As String
    Dim bShowCommand    As Boolean
    Dim sTempFile       As String
    Dim lResult         As Long
    
    sCommand = Command$()
    For Each vElem In SplitArgs(sCommand)
        Call OutputDebugString("Main: vElem=" & vElem & vbCrLf)
        Select Case Left$(vElem, 1)
        Case "-", "/"
            If LCase$(Mid(vElem, 2)) = "nologo" Then
                bShowCommand = True
            End If
        Case Else
            lPos = InStrRev(vElem, ".")
            If lPos > InStrRev(vElem, "\") Then
                sFile = Left$(vElem, lPos) & STR_COBJ_EXT
                If GetFileAttributes(sFile) <> INVALID_FILE_ATTRIBUTES Then
                    Call OutputDebugString("Main: sFile=" & sFile & vbCrLf)
                    sCommand = Replace(sCommand, vElem, sFile)
                End If
            End If
        End Select
    Next
    sCommand = """" & PathCombine(App.Path, STR_ORIGINAL_LINKER) & """ " & sCommand
    If bShowCommand Then
        MsgBox sCommand, vbExclamation
    End If
    sTempFile = Environ$("TEMP") & "\~$link.out"
    Call DeleteFile(sTempFile)
    lResult = ShellWait("cmd /c """ & sCommand & """ > " & sTempFile & " 2>&1", vbHide)
    With CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
        .Write ReadFile(sTempFile)
    End With
    Call ExitProcess(lResult)
End Sub
