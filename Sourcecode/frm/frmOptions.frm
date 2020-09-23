VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCount 
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   1755
      Width           =   960
   End
   Begin AppSheriff.dFlatButton CmdCancel 
      Height          =   375
      Left            =   3525
      TabIndex        =   5
      Top             =   2070
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
   End
   Begin VB.TextBox TxtB 
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   1080
      Width           =   4170
   End
   Begin VB.TextBox TxtA 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   405
      Width           =   4170
   End
   Begin AppSheriff.dFlatButton CmdOk 
      Height          =   375
      Left            =   2385
      TabIndex        =   4
      Top             =   2070
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OK"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowed number of attempts:"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   1515
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incorrect password error message:"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   855
      Width           =   2430
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dialog Title / Caption:"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   165
      Width           =   1545
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
    Unload frmOptions
End Sub

Private Sub CmdOk_Click()
    t_DataHead.PwsMsg1 = TxtA.Text
    t_DataHead.PwsMsg2 = TxtB.Text
    t_DataHead.PwsTrys = Val(txtCount.Text)
    'Set to 1 if value is zero
    If (t_DataHead.PwsTrys = 0) Then t_DataHead.PwsTrys = 1
    Call CmdCancel_Click
End Sub

Private Sub Form_Load()
    TxtA.Text = t_DataHead.PwsMsg1
    TxtB.Text = t_DataHead.PwsMsg2
    txtCount.Text = t_DataHead.PwsTrys
    Call FlatTxtBox(frmOptions)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub

Private Sub txtCount_KeyPress(KeyAscii As Integer)
    'Allow only digits.
    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
