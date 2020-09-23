Attribute VB_Name = "common"
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SND_FILENAME = &H20000
Public Const SND_ASYNC = &H1

Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_STATICEDGE As Long = &H20000

Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2

Private Type DataHead
    SIG As String * 5
    isEncrypted As Boolean
    PwsHash As Long
    ProgData() As Byte
    PwsMsg1 As String
    PwsMsg2 As String
    PwsTrys As Integer
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public t_DataHead As DataHead

Public Sub PlaySnd()
Dim sFile As String
    'Just plays the start.wav in Windows\media
    sFile = FixPath(Environ("WINDIR")) & "Media\start.wav"
    'Check if file is found.
    If IsFileHere(sFile) Then
        'Play the file.
        PlaySound sFile, App.hInstance, SND_FILENAME Or SND_ASYNC
    End If
    
    sFile = vbNullString
    
End Sub

Public Sub BevelPicBox(PicBox As PictureBox, BkColor As Long)
Dim rc As RECT
    FlatBorder PicBox.hwnd, True
    With PicBox
        rc.Left = 0
        rc.Right = 0
        rc.Right = .ScaleWidth - 1
        rc.Bottom = .ScaleHeight - 1
    
        .BackColor = BkColor
        DrawEdge .hDC, rc, &H4, &HF
        .Refresh
    End With
End Sub
Public Sub SetDataHead()
    With t_DataHead
        .PwsMsg1 = "Please enter password"
        .PwsMsg2 = "The password entered is incorrect. Please try again."
        .PwsTrys = 3
    End With
End Sub

Public Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
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

Public Function IsFileHere(ByVal lzFilename As String) As Boolean
    If Len(lzFilename) = 0 Then
        IsFileHere = False
        Exit Function
    End If
    
    If LenB(Dir(lzFilename)) = 0 Then
        IsFileHere = False
        Exit Function
    Else
        IsFileHere = True
    End If
    
End Function

Public Function IsFileOpen(lzFilename As String) As Boolean
Dim fp As Long
On Error GoTo ErrFlag:
    'Checks if a file is open.
    fp = FreeFile
    
    Open lzFilename For Binary Access Write As #fp
    Close #fp
    
    IsFileOpen = False
    
    Exit Function
ErrFlag:
    'File is Open
    IsFileOpen = True
    Close #fp
    
End Function

Function OpenFile(lzFile As String) As Byte()
Dim fp As Long
Dim fBytes() As Byte
    fp = FreeFile
    
    Open lzFile For Binary As #fp
        If LOF(fp) = 0 Then
            Exit Function
        Else
            ReDim Preserve fBytes(0 To LOF(fp) - 1)
            Get #fp, , fBytes
        End If
    Close #fp
    
    OpenFile = fBytes
    Erase fBytes
    
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

Public Sub FlatTxtBox(frm As Form)
Dim c As Control
    'Makes all Textboxes flat on a given form.
    For Each c In frm
        If TypeName(c) = "TextBox" Then FlatBorder c.hwnd, True
    Next c
    
    Set c = Nothing
End Sub

Public Function FlatBorder(ByVal hwnd As Long, MakeControlFlat As Boolean)
Dim TFlat As Long
    'Make control flat.
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    If MakeControlFlat Then
        TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        TFlat = TFlat And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    End If
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Function

Public Function IsExeLocked(ExeFile As String) As Boolean
Dim fp As Long
Dim sSIG As String
Dim Offset As Long
On Error Resume Next

    sSIG = Space(5)
    fp = FreeFile
    'Check if the exe is already locked.
    Open ExeFile For Binary As #fp
        'we only want the header info
        Get #fp, LOF(fp) - 3, Offset
        If (Offset = 0) Then
            Close #fp
            Exit Function
        Else
            'Get SIG
            Get #fp, Offset, sSIG
            IsExeLocked = LCase(sSIG) = "dlock"
        End If
    Close #fp
    
    sSIG = vbNullString
    Offset = 0
    
End Function
