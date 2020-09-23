VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM AppSheriff"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmBuild.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicStatBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   426
      TabIndex        =   40
      Top             =   3795
      Width           =   6390
      Begin AppSheriff.dFlatButton cmdHelp 
         Height          =   360
         Left            =   60
         TabIndex        =   42
         ToolTipText     =   "Help"
         Top             =   45
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         Caption         =   ""
         Picture         =   "frmBuild.frx":0442
      End
      Begin VB.Label lblBar 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   555
         TabIndex        =   41
         Top             =   120
         Width           =   90
      End
   End
   Begin VB.PictureBox PicTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   1320
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   336
      TabIndex        =   22
      Top             =   15
      Width           =   5040
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   75
         TabIndex        =   23
         Top             =   90
         Width           =   105
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   1320
      TabIndex        =   5
      Top             =   345
      Width           =   5040
      Begin VB.PictureBox pTab 
         BorderStyle     =   0  'None
         Height          =   2940
         Index           =   1
         Left            =   60
         ScaleHeight     =   2940
         ScaleWidth      =   4785
         TabIndex        =   30
         Top             =   135
         Width           =   4785
         Begin VB.TextBox txtPws2 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   150
            PasswordChar    =   "*"
            TabIndex        =   35
            Top             =   1095
            Width           =   3855
         End
         Begin VB.TextBox TxtFile2 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   435
            Width           =   3855
         End
         Begin VB.CheckBox chkBack 
            Caption         =   "Create backup of Locked Application."
            Height          =   240
            Left            =   135
            TabIndex        =   31
            Top             =   1530
            Width           =   4455
         End
         Begin AppSheriff.dFlatButton cmdUnlock 
            Height          =   350
            Left            =   3270
            TabIndex        =   32
            Top             =   2160
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            Caption         =   "UnLock"
            Enabled         =   0   'False
         End
         Begin AppSheriff.dFlatButton cmdOpen2 
            Height          =   300
            Left            =   4035
            TabIndex        =   33
            ToolTipText     =   "Select Application"
            Top             =   435
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            Caption         =   "...."
            ButtonStyle     =   2
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Application:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   37
            Top             =   135
            Width           =   1590
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Password:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   36
            Top             =   840
            Width           =   1395
         End
      End
      Begin VB.PictureBox pTab 
         BorderStyle     =   0  'None
         Height          =   2940
         Index           =   0
         Left            =   60
         ScaleHeight     =   2940
         ScaleWidth      =   4785
         TabIndex        =   6
         Top             =   135
         Width           =   4785
         Begin VB.CheckBox Chkbackup 
            Caption         =   "Backup Original Application."
            Height          =   240
            Left            =   135
            TabIndex        =   29
            Top             =   1815
            Width           =   4530
         End
         Begin VB.CheckBox ChkEncrypt 
            Caption         =   "Encrypt Application."
            Height          =   240
            Left            =   135
            TabIndex        =   27
            Top             =   1530
            Width           =   4515
         End
         Begin AppSheriff.dFlatButton CmdLock1 
            Height          =   345
            Left            =   3270
            TabIndex        =   26
            Top             =   2160
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            Caption         =   "Lock"
            Enabled         =   0   'False
         End
         Begin AppSheriff.dFlatButton cmdOpen1 
            Height          =   300
            Left            =   4035
            TabIndex        =   25
            ToolTipText     =   "Select Application"
            Top             =   435
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            Caption         =   "...."
            ButtonStyle     =   2
         End
         Begin VB.TextBox TxtFile1 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   435
            Width           =   3855
         End
         Begin VB.TextBox txtPws1 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   150
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   1095
            Width           =   3855
         End
         Begin AppSheriff.dFlatButton CmdOptions 
            Height          =   315
            Left            =   4035
            TabIndex        =   28
            ToolTipText     =   "Options"
            Top             =   1080
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   99
            Caption         =   ""
            Picture         =   "frmBuild.frx":04E9
            ButtonStyle     =   2
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Password:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   8
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblExeName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Application:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   7
            Top             =   135
            Width           =   1590
         End
      End
      Begin VB.PictureBox pTab 
         BorderStyle     =   0  'None
         Height          =   2940
         Index           =   2
         Left            =   60
         ScaleHeight     =   2940
         ScaleWidth      =   4770
         TabIndex        =   12
         Top             =   135
         Width           =   4770
         Begin VB.Label lblCpy 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright © 2004-2007 Ben Jones"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   150
            TabIndex        =   39
            Top             =   2580
            Width           =   2955
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808080&
            Height          =   735
            Left            =   180
            Top             =   1245
            Width           =   4440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "The fast, simple and easy way to protect your applications."
            Height          =   195
            Left            =   225
            TabIndex        =   38
            Top             =   930
            Width           =   4125
         End
         Begin VB.Image ImgLogo 
            Height          =   480
            Left            =   120
            Top             =   105
            Width           =   435
         End
         Begin VB.Label lblDesign 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Designed by DreamVB"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2625
            TabIndex        =   19
            Top             =   2070
            Width           =   1920
         End
         Begin VB.Label lblUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#0"
            Height          =   195
            Left            =   1470
            TabIndex        =   18
            Top             =   1650
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Licensed to:"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   17
            Top             =   1650
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Freeware"
            Height          =   195
            Index           =   4
            Left            =   1470
            TabIndex        =   16
            Top             =   1350
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "License Type:"
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   15
            Top             =   1350
            Width           =   1005
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   180
            X2              =   4605
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblTitle2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Version 1.2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3285
            TabIndex        =   14
            Top             =   525
            Width           =   975
         End
         Begin VB.Label lblTitle1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DM AppSheriff"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   705
            TabIndex        =   13
            Top             =   150
            Width           =   2025
         End
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   285
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3675
      Left            =   45
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   45
      Width           =   1245
      Begin VB.PictureBox pButton 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Index           =   3
         Left            =   15
         MousePointer    =   99  'Custom
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   20
         Top             =   2355
         Width           =   1155
         Begin VB.Image ImgButton 
            Height          =   480
            Index           =   3
            Left            =   285
            Picture         =   "frmBuild.frx":0646
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblButton 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   360
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   450
            Width           =   315
         End
      End
      Begin VB.PictureBox pButton 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Index           =   2
         Left            =   15
         MousePointer    =   99  'Custom
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   10
         Top             =   1575
         Width           =   1155
         Begin VB.Label lblButton 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   285
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   495
            Width           =   495
         End
         Begin VB.Image ImgButton 
            Height          =   480
            Index           =   2
            Left            =   285
            Picture         =   "frmBuild.frx":09B1
            Tag             =   "About"
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.PictureBox pButton 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Index           =   1
         Left            =   15
         MousePointer    =   99  'Custom
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   3
         Top             =   795
         Width           =   1155
         Begin VB.Image ImgButton 
            Height          =   480
            Index           =   1
            Left            =   300
            Picture         =   "frmBuild.frx":0F83
            Tag             =   "UnLock Application"
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblButton 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unlock"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   270
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   525
            Width           =   570
         End
      End
      Begin VB.PictureBox pButton 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Index           =   0
         Left            =   15
         MouseIcon       =   "frmBuild.frx":1206
         MousePointer    =   99  'Custom
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   1
         Top             =   15
         Width           =   1155
         Begin VB.Label lblButton 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lock"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   345
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   525
            Width           =   390
         End
         Begin VB.Image ImgButton 
            Height          =   480
            Index           =   0
            Left            =   270
            Picture         =   "frmBuild.frx":1358
            Tag             =   "Lock Application"
            Top             =   0
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private exe_head As String
Private m_last As Integer

Private Function OpenDLG(Optional Title As String = "Open") As String
On Error GoTo OpenErr:

    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = "Program Files(*.exe)|*.exe|"
        .ShowOpen
        'Check for Exe File.
        If UCase(Right(.Filename, 3)) <> "EXE" Then
            MsgBox "File not supported.", vbCritical, "File Format Not Supported"
            Exit Function
        Else
            'Return Filename.
            OpenDLG = .Filename
        End If
    End With
    
    Exit Function
OpenErr:
    If Err Then Err.Clear
    
End Function

Private Sub cmdHelp_Click()
    ShellExecute frmmain.hwnd, "open", FixPath(App.Path) & "help.chm", vbNullString, vbNullString, 1
End Sub

Private Sub cmdOpen1_Click()
Dim Filename As String
    'Get Filename from Dialog.
    Filename = OpenDLG("Select Application")
    
    If Len(Filename) > 0 Then
        'Check if the exe is not already locked
        If IsExeLocked(Filename) Then
            MsgBox "This application has already been locked.", vbCritical, _
            "File Already Locked"
            Exit Sub
        Else
            'Update textbox with Filename.
            TxtFile1.Text = Filename
            TxtFile1.ToolTipText = Filename
            Filename = vbNullString
        End If
    End If
End Sub

Private Sub cmdOpen2_Click()
Dim Filename As String
    'Get Filename from Dialog.
    Filename = OpenDLG("Select Application")
    
    If Len(Filename) > 0 Then
        'Check if the exe has been locked.
        If Not IsExeLocked(Filename) Then
            MsgBox "This application has not been locked by " & frmmain.Caption & _
            " Or the file has already been unlocked.", vbCritical, "File Not Locked"
            Exit Sub
        Else
            'Update textbox with Filename.
            TxtFile2.Text = Filename
            TxtFile2.ToolTipText = Filename
            Filename = vbNullString
        End If
    End If
End Sub

Private Sub CmdUnlock_Click()
Dim PwsA As String
Dim SrcFile As String
Dim Offset As Long
Dim fp As Long
    
    SrcFile = Trim(TxtFile2.Text)
    'Store the Password hash.
    PwsA = txtPws2.Text
    'Check if the exe to be protected exsists.
    If Not IsFileHere(SrcFile) Then
        MsgBox "The locked application was not found." & vbCrLf & SrcFile, vbExclamation, "File Not Found"
        Exit Sub
    End If
    'Check that the exe is not already open
    If IsFileOpen(SrcFile) Then
        MsgBox "Error UnLocking File:" & vbCrLf & SrcFile, vbCritical, "Access denied"
        Exit Sub
    End If
    'Check if the exe it not already locked.
    If Not IsExeLocked(SrcFile) Then
        MsgBox "This application has not been locked by " & frmmain.Caption & _
        " or the file has already been unlocked.", vbCritical, "File Not Locked"
        Exit Sub
    End If
    'Open the locked exe and extract the data
    fp = FreeFile
    Open SrcFile For Binary As #fp
        Get #fp, LOF(fp) - 3, Offset
        'Get the info header
        Get #fp, Offset, t_DataHead
    Close #fp
    'Check for correct password.
    If MakePwsHash(PwsA) <> t_DataHead.PwsHash Then
        MsgBox "The password entered was incorrect." _
        & vbCrLf & "Please try entering the password again.", vbExclamation, "Wrong Password"
        Exit Sub
    End If
    'Check if the file has been encrypted.
    If (t_DataHead.isEncrypted) Then Call ByteXorClipper(t_DataHead.ProgData, PwsA)
    'Check if backup is enabled.
    If (chkBack) Then FileCopy SrcFile, SrcFile & ".bak"
    'Delete the original exe
    If Not DoKillFile(SrcFile) Then
        MsgBox "Error UnLocking File:" & vbCrLf & SrcFile, vbCritical, "Access denied"
        t_DataHead.isEncrypted = False
        t_DataHead.PwsHash = 0
        Erase t_DataHead.ProgData
        Exit Sub
    Else
        'Now we restore the original exe
        Open SrcFile For Binary As #fp
            Put #fp, , t_DataHead.ProgData
        Close #fp
    End If
    
    MsgBox "Your application has now successfully been unlocked by " & _
    frmmain.Caption, vbInformation, "Finished"
    
    'Erase file data
    Offset = 0
    SrcFile = vbNullString
    PwsA = vbNullString
    Erase t_DataHead.ProgData
    t_DataHead.PwsHash = 0
    t_DataHead.isEncrypted = False
    
End Sub

Private Sub CmdLock1_Click()
Dim PwsA As String
Dim SrcFile As String
Dim Offset As Long
Dim fp As Long
    
    SrcFile = Trim(TxtFile1.Text)
    'Store the Password hash.
    PwsA = txtPws1.Text
    
    'Check if the exe to be protected exsists.
    If Not IsFileHere(SrcFile) Then
        MsgBox "The application was not found." & vbCrLf & SrcFile, vbExclamation, "File Not Found"
        Exit Sub
    End If
    'Check if the exe it not already locked.
    If IsExeLocked(SrcFile) Then
        MsgBox "This application has already been locked.", vbCritical, _
        "File Already Locked"
        Exit Sub
    End If
    'Fill in t_DataHead info
    With t_DataHead
        .SIG = "dLock"
        .isEncrypted = ChkEncrypt
        .ProgData = OpenFile(SrcFile)
        .PwsHash = MakePwsHash(PwsA)
        'Check if encryption is enabled.
        If (ChkEncrypt) Then Call ByteXorClipper(.ProgData, PwsA)
    End With
    
    'Check if backup is enabled
    If (Chkbackup) Then FileCopy SrcFile, SrcFile & ".bak"
    
    'Make sure original exe is not already open
    If IsFileOpen(SrcFile) Then
        MsgBox "Error Locking File:" & vbCrLf & SrcFile, vbCritical, "Access denied"
        Exit Sub
    'Delete original exe
    ElseIf Not DoKillFile(SrcFile) Then
        MsgBox "Error Locking File:" & vbCrLf & SrcFile, vbCritical, "Access denied"
    Else
        'Copy stub file over as original exe name
        FileCopy exe_head, SrcFile
        'Merge t_DataHead into SrcFile
        fp = FreeFile
        Open SrcFile For Binary As #fp
            'Move to the end of the file
            'Get FileLen
            Offset = LOF(fp)
            'Seek to End of file
            Seek #fp, Offset
            'Write t_DataHead
            Put #fp, , t_DataHead
            'Write offset
            Put #fp, , Offset
        Close #fp
    End If
    
    MsgBox "Your application has now successfully been locked by " & _
    frmmain.Caption, vbInformation, "Finished"
    
    'Erase file data
    Offset = 0
    SrcFile = vbNullString
    PwsA = vbNullString
    Erase t_DataHead.ProgData
    t_DataHead.PwsHash = 0
    t_DataHead.isEncrypted = False
End Sub

Private Sub CmdOptions_Click()
    'Show password options dialog.
    frmOptions.Show vbModal, frmmain
End Sub

Private Sub Form_Load()
Dim c As Control

    exe_head = FixPath(App.Path) & "stub.dat"
    
    'Check if exe_head was found.
    If Not IsFileHere(exe_head) Then
        MsgBox exe_head & vbCrLf _
        & "Was not found on the system. Please try reinstalling the program again.", vbCritical, "File Not Found"
        Unload frmmain
        Exit Sub
    End If
    
    lblUser.Caption = Environ("USERNAME")
    lblBar.Caption = frmmain.Caption & " Version 1.2"
    'Sets the hand cursor for the controls
    For Each c In frmmain
        If TypeName(c) = "dFlatButton" Then
            Set c.MouseIcon = pButton(0).MouseIcon
        End If
        
        If (c.Name = "pButton") Then
            If (c.Index <> 0) Then
                Set c.MouseIcon = pButton(0).MouseIcon
            End If
        End If
    Next c
    Set c = Nothing
    
    'Draws Bevel on Pixbox
    Call BevelPicBox(PicTitle, &HC56A31)
    Call BevelPicBox(PicStatBar, vbButtonFace)
    'Make all textboxes flat.
    Call FlatTxtBox(frmmain)
    'Logo for about box
    ImgLogo.Picture = ImgButton(0).Picture
    'Selects first item
    Call pButton_MouseDown(0, 1, 0, 0, 0)
    Call pButton_MouseUp(0, 1, 0, 0, 0)
    
    Call SetDataHead
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub ImgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pButton_MouseMove Index, Button, Shift, 0, 0
End Sub

Private Sub lblButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pButton_MouseMove Index, Button, Shift, 0, 0
End Sub

Private Sub pButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    Call ButtonSelected(Index)
End Sub

Private Sub ButtonSelected(Index As Integer)
Dim X As Integer
    'Used for selected button style
    For X = 0 To pButton.Count - 1
        pButton(X).Tag = vbNullString
        pButton(X).BackColor = vbWhite
    Next X
    'Set button tag to selected.
    pButton(Index).Tag = 1
    pButton_MouseMove Index, 1, 0, 0, 0
End Sub

Private Sub pButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With pButton(Index)
        
        'Check if the button is selected.
        If (.Tag = "1") Then
            .BackColor = &HEED2C1
            pButton(Index).Line (0, 0)-(pButton(Index).ScaleWidth - 1, pButton(Index).ScaleHeight - 1), &HC56A31, B
            Exit Sub
        End If
        
        'Do Hover effects.
        If (X < 0) Or (X > .ScaleWidth) Or (Y < 0) Or (Y > .ScaleHeight) Then
            ReleaseCapture
            .BackColor = vbWhite
        ElseIf GetCapture() <> .hwnd Then
            SetCapture .hwnd
            .BackColor = &HF6E8E0
            pButton(Index).Line (0, 0)-(pButton(Index).ScaleWidth - 1, pButton(Index).ScaleHeight - 1), &HE2B498, B
        End If
 
    End With
End Sub

Private Sub pButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    If (Index = 3) Then
        If MsgBox("Do you want to exit the program now?", vbYesNo Or vbQuestion, frmmain.Caption) = vbYes Then
            Unload frmmain
        Else
            ButtonSelected m_last
        End If
    Else
        lblCaption.Caption = ImgButton(Index).Tag
        ArrangeTabs Index
    End If
    
    If (Index <> 3) Then m_last = Index
End Sub

Private Sub ArrangeTabs(Index As Integer)
Dim Count As Long

    For Count = 0 To pTab.Count - 1
        pTab(Count).Visible = False
    Next Count
    
    pTab(Index).Visible = True
    
    Call PlaySnd
End Sub

Private Sub TxtFile1_Change()
    txtPws1_Change
End Sub

Private Sub TxtFile2_Change()
    txtPws2_Change
End Sub

Private Sub txtPws1_Change()
    CmdLock1.Enabled = Len(Trim(TxtFile1.Text)) > 0 And Len(Trim(txtPws1.Text)) > 0
End Sub

Private Sub txtPws2_Change()
    cmdUnlock.Enabled = Len(Trim(TxtFile2.Text)) > 0 And Len(Trim(txtPws2.Text)) > 0
End Sub
