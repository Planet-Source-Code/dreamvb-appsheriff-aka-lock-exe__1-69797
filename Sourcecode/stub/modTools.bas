Attribute VB_Name = "modTools"
Option Explicit

Private Const WAIT_INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

Private Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessID As Long) As Long
   
Private Declare Function WaitForSingleObject Lib "kernel32" _
  (ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long
   
Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long


Private Type DataHead
    SIG As String * 5
    isEncrypted As Boolean
    PwsHash As Long
    ProgData() As Byte
    PwsMsg1 As String
    PwsMsg2 As String
    PwsTrys As Integer
End Type

Public t_DataHead As DataHead

Public Sub CleanDataHead()
    With t_DataHead
        Erase .ProgData
        .isEncrypted = False
        .PwsHash = 0
        .PwsMsg1 = vbNullString
        .PwsMsg2 = vbNullString
        .SIG = vbNullString
        .PwsTrys = 0
    End With
End Sub

Public Function VbShellWait(lzApp As String, FrmObj As Form) As Integer
Dim hProcess As Long
Dim ProcessID As Long
Dim Lock_FilePtr As Long
On Error GoTo LoadErr:
    
    VbShellWait = 0
    ProcessID = Shell(lzApp, vbNormalFocus)
    hProcess = OpenProcess(SYNCHRONIZE, True, ProcessID)
    'Lock the file
    Lock_FilePtr = FreeFile
    Open lzApp For Binary Lock Read As #Lock_FilePtr
        'Hide form from user
        FrmObj.Hide
        Call WaitForSingleObject(hProcess, WAIT_INFINITE)
       ' RegisterServiceProcess hProcess, 1
        CloseHandle hProcess
        VbShellWait = 1
    Close #Lock_FilePtr
    Exit Function
LoadErr:
    Close #Lock_FilePtr
    VbShellWait = 2
    
End Function

Function FixPath(lPath As String) As String
    If Right(lPath, 1) <> "\" Then
        FixPath = lPath & "\"
    Else
        FixPath = lPath
    End If
End Function

Public Function DoKillFile(lzFile As String) As Boolean
    On Error GoTo ErrFlag:
    
    SetAttr lzFile, vbNormal
    Kill lzFile
    DoKillFile = True
    
    Exit Function
ErrFlag:
    If Err Then
        DoKillFile = False
    End If
    
End Function

Public Function MakePwsHash(StrPws As String) As Long
Dim X As Long
Dim xCode As Long
Dim c As Long

    c = 128
    For X = 1 To Len(StrPws)
        c = Asc(Mid(StrPws, X, 1)) + c
    Next X

    MakePwsHash = c
    c = 0
End Function

Public Sub ByteXorClipper(bBytes() As Byte, Pws As String)
Dim X As Long
Dim PwsArr() As Byte
Dim PwsSize As Long
Dim idx As Integer
    
    PwsArr = StrConv(Pws, vbFromUnicode)
    PwsSize = UBound(PwsArr)
    Rnd (-3)
    
    For X = 0 To UBound(bBytes)
        bBytes(X) = bBytes(X) Xor PwsArr(idx) Xor Int(255 * Rnd)
        If (idx >= PwsSize) Then idx = -1
        idx = (idx + 1)
    Next X
    
    Erase PwsArr
    X = 0
    PwsSize = 0
End Sub

