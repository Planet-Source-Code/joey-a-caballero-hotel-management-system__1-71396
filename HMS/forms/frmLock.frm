VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox bgPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   30
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   1
      Top             =   4500
      Width           =   6435
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   60
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1230
         TabIndex        =   2
         Top             =   90
         Width           =   1605
      End
   End
   Begin VB.PictureBox imgBg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   30
      ScaleHeight     =   4425
      ScaleWidth      =   6435
      TabIndex        =   3
      Top             =   30
      Width           =   6435
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   -360
         ScaleHeight     =   645
         ScaleWidth      =   7005
         TabIndex        =   4
         Top             =   0
         Width           =   7035
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Locked"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   735
            Left            =   450
            TabIndex        =   5
            Top             =   -90
            Width           =   2145
         End
      End
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private br As RECT


Public Function ShowForm()
On Error Resume Next
    Me.Show vbModal
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub



Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
    If lpass = txtPassword.Text Then
        Unload Me
        Exit Sub
    End If
    
    If KeyCode = vbKeyEscape Then
        bgPassword.Visible = False
    End If
        
End Sub
Private Sub Form_Load()
    'Load frmSplash
    'Set imgBg.Picture = frmSplash.Picture
    'Unload frmSplash
End Sub
