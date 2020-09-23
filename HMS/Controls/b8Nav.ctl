VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl b8Nav 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   540
      Top             =   1650
   End
   Begin VB.PictureBox imgNav 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1380
      Picture         =   "b8Nav.ctx":0000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   3
      ToolTipText     =   "Display Last List"
      Top             =   0
      Width           =   345
   End
   Begin VB.PictureBox imgNav 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1005
      Picture         =   "b8Nav.ctx":0471
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   2
      ToolTipText     =   "Display Next List"
      Top             =   0
      Width           =   345
   End
   Begin VB.PictureBox imgNav 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   510
      Picture         =   "b8Nav.ctx":08BB
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   1
      ToolTipText     =   "Display Previous List"
      Top             =   0
      Width           =   345
   End
   Begin VB.PictureBox imgNav 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   150
      Picture         =   "b8Nav.ctx":0D13
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   0
      ToolTipText     =   "Display First List"
      Top             =   0
      Width           =   345
   End
   Begin MSComctlLib.ImageList ilDis 
      Left            =   1230
      Top             =   795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":19B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":1DDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilHot 
      Left            =   2070
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":2203
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":26B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":2B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":2FB9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilNormal 
      Left            =   315
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":345B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":38EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":3D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "b8Nav.ctx":41AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgBG 
      Height          =   375
      Left            =   0
      Picture         =   "b8Nav.ctx":462F
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "b8Nav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Dim MouseOnDown As Boolean
Dim CurIndex As Integer

'Default Property Values:
Const m_def_FirstEnable = True
Const m_def_PreviousEnable = True
Const m_def_NextEnable = True
Const m_def_LastEnable = True
'Property Variables:
Dim m_FirstEnable As Boolean
Dim m_PreviousEnable As Boolean
Dim m_NextEnable As Boolean
Dim m_LastEnable As Boolean
'Event Declarations:
Event Click(Index As Integer)




Private Sub imgNav_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgNav(Index).Picture = IIf(imgNav(Index).Enabled, ilNormal.ListImages(Index + 1).Picture, ilDis.ListImages(Index + 1).Picture)
    
    MouseOnDown = True
End Sub

Private Sub imgNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    For i = 0 To imgNav.UBound
        If Index <> i Then
            If imgNav(i).Enabled = True Then
                imgNav(i).Picture = ilNormal.ListImages(i + 1).Picture
            Else
                imgNav(i).Picture = ilDis.ListImages(i + 1).Picture

            End If
        End If
    Next
    
    If imgNav(Index).Enabled = False Then Exit Sub
    
    If MouseOnDown = False Then
        imgNav(Index).Picture = ilHot.ListImages(Index + 1).Picture
    End If
    
    CurIndex = Index
    timerMouse.Enabled = True
End Sub

Private Sub imgNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim P As POINTAPI
    Dim R As RECT
    
    imgNav(Index).Picture = IIf(imgNav(Index).Enabled, ilNormal.ListImages(Index + 1).Picture, ilDis.ListImages(Index + 1).Picture)
    
    GetWindowRect imgNav(Index).hwnd, R
    GetCursorPos P
    
    If Not (P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom) Then
        
        RaiseEvent Click(Index)
        
    End If
    
    
    'RaiseEvent CloseMouseUp(Button, Shift, X, Y)
    MouseOnDown = False
End Sub

Private Sub timerMouse_Timer()
    Dim P As POINTAPI
    Dim R As RECT

    GetWindowRect imgNav(CurIndex).hwnd, R
    GetCursorPos P
    
    If P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom Then
        timerMouse.Enabled = False
        If imgNav(CurIndex).Enabled = True Then
            imgNav(CurIndex).Picture = ilNormal.ListImages(CurIndex + 1).Picture
        Else
            imgNav(CurIndex).Picture = ilDis.ListImages(CurIndex + 1).Picture
        End If
    End If
End Sub



'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    imgNav(0).Enabled = PropBag.ReadProperty("FirstEnable", True)
    imgNav(1).Enabled = PropBag.ReadProperty("PreviousEnable", True)
    imgNav(2).Enabled = PropBag.ReadProperty("NextEnable", True)
    imgNav(3).Enabled = PropBag.ReadProperty("LastEnable", True)
    m_FirstEnable = PropBag.ReadProperty("FirstEnable", m_def_FirstEnable)
    m_PreviousEnable = PropBag.ReadProperty("PreviousEnable", m_def_PreviousEnable)
    m_NextEnable = PropBag.ReadProperty("NextEnable", m_def_NextEnable)
    m_LastEnable = PropBag.ReadProperty("LastEnable", m_def_LastEnable)
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = vbPixels
    UserControl.Width = imgBG.Width * Screen.TwipsPerPixelX
    UserControl.Height = imgBG.Height * Screen.TwipsPerPixelY

End Sub
Public Sub RefreshButtons()
    Dim i As Integer
    
    For i = 0 To imgNav.UBound

            If imgNav(i).Enabled = True Then
                imgNav(i).Picture = ilNormal.ListImages(i + 1).Picture
                
            Else
                imgNav(i).Picture = ilDis.ListImages(i + 1).Picture
            End If
    Next
    
    
End Sub

Private Sub UserControl_Show()
    
    RefreshButtons
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("FirstEnable", imgNav(0).Enabled, True)
    Call PropBag.WriteProperty("PreviousEnable", imgNav(1).Enabled, True)
    Call PropBag.WriteProperty("NextEnable", imgNav(2).Enabled, True)
    Call PropBag.WriteProperty("LastEnable", imgNav(3).Enabled, True)
    Call PropBag.WriteProperty("FirstEnable", m_FirstEnable, m_def_FirstEnable)
    Call PropBag.WriteProperty("PreviousEnable", m_PreviousEnable, m_def_PreviousEnable)
    Call PropBag.WriteProperty("NextEnable", m_NextEnable, m_def_NextEnable)
    Call PropBag.WriteProperty("LastEnable", m_LastEnable, m_def_LastEnable)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get FirstEnable() As Boolean
Attribute FirstEnable.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    FirstEnable = m_FirstEnable
End Property

Public Property Let FirstEnable(ByVal New_FirstEnable As Boolean)
    m_FirstEnable = New_FirstEnable
    imgNav(0).Enabled = New_FirstEnable
    PropertyChanged "FirstEnable"
    RefreshButtons
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get PreviousEnable() As Boolean
Attribute PreviousEnable.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    PreviousEnable = m_PreviousEnable
End Property

Public Property Let PreviousEnable(ByVal New_PreviousEnable As Boolean)
    m_PreviousEnable = New_PreviousEnable
    imgNav(1).Enabled = New_PreviousEnable
    PropertyChanged "PreviousEnable"
    RefreshButtons
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get NextEnable() As Boolean
Attribute NextEnable.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    NextEnable = m_NextEnable
End Property

Public Property Let NextEnable(ByVal New_NextEnable As Boolean)
    m_NextEnable = New_NextEnable
    imgNav(2).Enabled = New_NextEnable
    PropertyChanged "NextEnable"
    RefreshButtons
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get LastEnable() As Boolean
Attribute LastEnable.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    LastEnable = m_LastEnable
End Property

Public Property Let LastEnable(ByVal New_LastEnable As Boolean)
    m_LastEnable = New_LastEnable
    imgNav(3).Enabled = New_LastEnable
    PropertyChanged "LastEnable"
    RefreshButtons
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FirstEnable = m_def_FirstEnable
    m_PreviousEnable = m_def_PreviousEnable
    m_NextEnable = m_def_NextEnable
    m_LastEnable = m_def_LastEnable
End Sub

