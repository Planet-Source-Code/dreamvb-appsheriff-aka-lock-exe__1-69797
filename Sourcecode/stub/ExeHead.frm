VERSION 5.00
Begin VB.Form frmExeHead 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "ExeHead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Top             =   990
      Width           =   1060
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   1830
      TabIndex        =   2
      Top             =   990
      Width           =   1060
   End
   Begin VB.TextBox txtA10 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   330
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   3675
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#0"
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
      Left            =   360
      TabIndex        =   0
      Top             =   225
      Width           =   240
   End
End
Attribute VB_Name = "frmExeHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Prog_File As String

Private Sub CleanUp()
On Error Resume Next
    Call CleanDataHead
    Prog_File = ""
    Unload frmExeHead
End Sub

Private Sub cmdCancel_Click()
    Call CleanUp
End Sub

Private Sub cmdOK_Click()
Dim NewName As String
Dim A154 As String
Dim fp As Long
Static iCount As Integer
    
    App.TaskVisible = False
    App.Title = ""

    A154 = txtA10.Text 'password.
    'Check password hash match
    If (MakePwsHash(A154) <> t_DataHead.PwsHash) Then
        'Error message to display for incorrect password.
        MsgBox t_DataHead.PwsMsg2, vbExclamation, "Error_52"
        'Inc number of times tryed.
        iCount = (iCount + 1)
        'Check max try attempts.
        If (iCount >= t_DataHead.PwsTrys) Then
            'Max reached exit.
            Call CleanUp
            'Exit Sub
        End If
        Exit Sub
    Else
        'Correct password was entered.
        NewName = Prog_File & ".exe"
        'Check if the exe is encrypyed
        If (t_DataHead.isEncrypted) Then
            'Data is encrypted and needs decrypting
            Call ByteXorClipper(t_DataHead.ProgData, A154)
        End If
        'write the exe to temp file
        fp = FreeFile
        Open NewName For Binary As #fp
            Put #fp, , t_DataHead.ProgData
        Close #fp
        '
        Call CleanUp
        A154 = vbNullString
    End If
    'Execute the exe
    
    If VbShellWait(NewName, frmExeHead) <> 1 Then
        MsgBox "There was an error loading the application.", vbCritical, "Error_54"
        Call CleanUp
        Exit Sub
    Else
        DoKillFile NewName
        Call CleanUp
    End If
    
    'Clear up
    A154 = vbNullString
    NewName = vbNullString
    
End Sub

Private Sub Form_Load()
Dim fp As Long
Dim sPos As Long
Dim sSig As String

    Const errMsg1 = "Data Segment Read Error"
    Set frmExeHead.Icon = Nothing
    
    Prog_File = FixPath(App.Path) & App.EXEName & ".exe"
    '
    fp = FreeFile
    Open Prog_File For Binary As #fp
        'Get Data Start offset
        Get #fp, LOF(fp) - 3, sPos
        'Check for vaild offset
        If (sPos = 0) Then
            MsgBox errMsg1, vbCritical, "Error_53"
            Close #fp
            Unload frmExeHead
            Exit Sub
        Else
            sSig = Space(5)
            Get #fp, sPos, sSig
            'Check the sig is vaild
            If LCase(sSig) <> "dlock" Then
                MsgBox errMsg1, vbCritical, "Error_53"
                sSig = vbNullString
                Close #fp
            Else
                'Get the app data
                Get #fp, sPos, t_DataHead
            End If
        End If
    Close #fp
    'Update titles
    lblTitle.Caption = t_DataHead.PwsMsg1
    frmExeHead.Caption = lblTitle.Caption
    sSig = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmExeHead = Nothing
End Sub

