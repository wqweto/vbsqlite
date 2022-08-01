Attribute VB_Name = "mdMain"
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const INVALID_FILE_ATTRIBUTES       As Long = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_COBJ_EXT              As String = "cobj"
Private Const STR_ORIGINAL_LINKER       As String = "vblink.exe"

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Dim sCommand        As String
    Dim vElem           As Variant
    Dim lPos            As Long
    Dim sFile           As String
    Dim bShowCommand    As Boolean
    Dim sOutFile        As String
    Dim sTempFile       As String
    Dim lResult         As Long
    
    On Error GoTo EH
    sCommand = ArgvQuote(PathCombine(App.Path, STR_ORIGINAL_LINKER))
    For Each vElem In SplitArgs(Command$)
        Call OutputDebugString("Main: vElem=" & vElem & vbCrLf)
        Select Case Left$(vElem, 1)
        Case "-", "/"
            If LCase$(Mid(vElem, 2)) = "nologo" Then
                bShowCommand = True
            ElseIf LCase$(Mid$(vElem, 2, 4)) = "out:" Then
                sOutFile = Mid$(vElem, 6)
                Call OutputDebugString("Main: sOutFile=" & sOutFile & vbCrLf)
                ChDrive Left$(sOutFile, 2)
                ChDir Left$(sOutFile, InStrRev(sOutFile, "\"))
            End If
        Case Else
            lPos = InStrRev(vElem, ".")
            If lPos > InStrRev(vElem, "\") Then
                sFile = Left$(vElem, lPos) & STR_COBJ_EXT
                If GetFileAttributes(sFile) <> INVALID_FILE_ATTRIBUTES Then
                    Call OutputDebugString("Main: " & vElem & "->" & sFile & vbCrLf)
                    vElem = sFile
                End If
            End If
        End Select
        sCommand = sCommand & " " & ArgvQuote(CStr(vElem))
    Next
    If bShowCommand Then
        Clipboard.Clear
        Clipboard.SetText sCommand
        MsgBox "Command copied to clipboard: " & sCommand, vbExclamation
    End If
    sTempFile = Environ$("TEMP") & "\~$link.out"
    Call DeleteFile(sTempFile)
    lResult = ShellWait("cmd /c """ & sCommand & """ > " & ArgvQuote(sTempFile) & " 2>&1", vbHide)
    With CreateObject("Scripting.FileSystemObject").GetStandardStream(1)
        .Write ReadFile(sTempFile)
    End With
    Call ExitProcess(lResult)
    Exit Sub:
EH:
    Call OutputDebugString("Critical error: " & Err.Description & vbCrLf)
    Call ExitProcess(1)
End Sub

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
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
    If Left(sContents, 2) = Chr$(&HFF) & Chr$(&HFE) Then
        sContents = Mid(StrConv(sContents, vbFromUnicode), 2)
    End If
    ReadFile = sContents
QH:
End Function

' based on https://blogs.msdn.microsoft.com/twistylittlepassagesallalike/2011/04/23/everyone-quotes-command-line-arguments-the-wrong-way
Public Function ArgvQuote(sArg As String, Optional ByVal Force As Boolean) As String
    Const WHITESPACE As String = "*[ " & vbTab & vbVerticalTab & vbCrLf & "]*"
    
    If Not Force And LenB(sArg) <> 0 And Not sArg Like WHITESPACE Then
        ArgvQuote = sArg
    Else
        With CreateObject("VBScript.RegExp")
            .Global = True
            .Pattern = "(\\+)($|"")|(\\+)"
            ArgvQuote = """" & Replace(.Replace(sArg, "$1$1$2$3"), """", "\""") & """"
        End With
    End If
End Function

Public Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\", "\", vbNullString) & sFile
End Function

