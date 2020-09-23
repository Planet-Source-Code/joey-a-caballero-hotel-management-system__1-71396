VERSION 5.00
Begin VB.UserControl b8Navigator 
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   Begin VB.PictureBox cmdNormal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   1080
      Picture         =   "b8Navigator.ctx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   150
      Width           =   240
   End
   Begin VB.PictureBox cmdNormal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   810
      Picture         =   "b8Navigator.ctx":0222
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   150
      Width           =   240
   End
   Begin VB.PictureBox cmdNormal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   540
      Picture         =   "b8Navigator.ctx":0433
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   150
      Width           =   240
   End
   Begin VB.PictureBox cmdNormal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   270
      Picture         =   "b8Navigator.ctx":0644
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   150
      Width           =   240
   End
   Begin VB.Timer timerCmdOver 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2430
      Top             =   1410
   End
   Begin VB.Image cmdNoFocus 
      Height          =   240
      Index           =   3
      Left            =   1590
      Picture         =   "b8Navigator.ctx":0866
      Top             =   810
      Width           =   240
   End
   Begin VB.Image cmdNoFocus 
      Height          =   240
      Index           =   2
      Left            =   1380
      Picture         =   "b8Navigator.ctx":0A88
      Top             =   750
      Width           =   240
   End
   Begin VB.Image cmdNoFocus 
      Height          =   240
      Index           =   1
      Left            =   1020
      Picture         =   "b8Navigator.ctx":0C99
      Top             =   720
      Width           =   240
   End
   Begin VB.Image cmdNoFocus 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "b8Navigator.ctx":0EAA
      Top             =   810
      Width           =   240
   End
   Begin VB.Image cmdNavHL 
      Height          =   240
      Index           =   3
      Left            =   1530
      Picture         =   "b8Navigator.ctx":10CC
      Top             =   1290
      Width           =   240
   End
   Begin VB.Image cmdNavHL 
      Height          =   240
      Index           =   2
      Left            =   1260
      Picture         =   "b8Navigator.ctx":12F1
      Top             =   1290
      Width           =   240
   End
   Begin VB.Image cmdNavHL 
      Height          =   240
      Index           =   1
      Left            =   990
      Picture         =   "b8Navigator.ctx":1506
      Top             =   1290
      Width           =   240
   End
   Begin VB.Image cmdNavHL 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "b8Navigator.ctx":171E
      Top             =   1290
      Width           =   240
   End
   Begin VB.Image cmdNavDisabled 
      Height          =   240
      Index           =   3
      Left            =   1530
      Picture         =   "b8Navigator.ctx":1943
      Top             =   1020
      Width           =   240
   End
   Begin VB.Image cmdNavDisabled 
      Height          =   240
      Index           =   2
      Left            =   1260
      Picture         =   "b8Navigator.ctx":1B1B
      Top             =   1020
      Width           =   240
   End
   Begin VB.Image cmdNavDisabled 
      Height          =   240
      Index           =   1
      Left            =   990
      Picture         =   "b8Navigator.ctx":1CE0
      Top             =   1020
      Width           =   240
   End
   Begin VB.Image cmdNavDisabled 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "b8Navigator.ctx":1EA5
      Top             =   1020
      Width           =   240
   End
End
Attribute VB_Name = "b8Navigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim OnMouseOverIndex As Integer
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type


Private Sub cmdNormal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'While timerCmdOver.Enabled = True
    '    DoEvents
    'Wend
    
    OnMouseOverIndex = Index
    timerCmdOver.Enabled = True
    
End Sub

Private Sub timerCmdOver_Timer()
    Dim R As RECT
    Dim P As POINTAPI
    
    GetWindowRect cmdNormal(OnMouseOverIndex).hwnd, R
    GetCursorPos P
    
    If P.X >= R.Left And P.X <= R.Right And P.Y >= R.Top And P.Y <= R.Bottom Then
        'mouse over
        cmdNormal(OnMouseOverIndex).Picture = cmdNavHL(OnMouseOverIndex).Picture
    Else
        'mouse out
        cmdNormal(OnMouseOverIndex).Picture = cmdNormal(OnMouseOverIndex).Picture
        'timerCmdOver.Enabled = False
    End If
End Sub

