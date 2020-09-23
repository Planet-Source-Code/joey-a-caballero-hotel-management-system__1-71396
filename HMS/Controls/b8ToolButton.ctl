VERSION 5.00
Begin VB.UserControl b8ToolButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4410
      Top             =   1095
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   210
      Picture         =   "b8ToolButton.ctx":0000
      Top             =   1350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image i2 
      Height          =   690
      Index           =   1
      Left            =   4020
      Picture         =   "b8ToolButton.ctx":08CA
      Stretch         =   -1  'True
      Top             =   90
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image i3 
      Height          =   690
      Index           =   1
      Left            =   5340
      Picture         =   "b8ToolButton.ctx":09C4
      Stretch         =   -1  'True
      Top             =   210
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image i1 
      Height          =   690
      Index           =   1
      Left            =   3690
      Picture         =   "b8ToolButton.ctx":0C2E
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image i2 
      Height          =   690
      Index           =   0
      Left            =   2640
      Picture         =   "b8ToolButton.ctx":0E98
      Stretch         =   -1  'True
      Top             =   780
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image i3 
      Height          =   690
      Index           =   0
      Left            =   3960
      Picture         =   "b8ToolButton.ctx":0F92
      Stretch         =   -1  'True
      Top             =   900
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image i1 
      Height          =   690
      Index           =   0
      Left            =   2310
      Picture         =   "b8ToolButton.ctx":11FC
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgIcon2 
      Height          =   585
      Left            =   3660
      Top             =   1740
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "b8ToolButton"
      ForeColor       =   &H00BBCACD&
      Height          =   195
      Left            =   930
      TabIndex        =   0
      Top             =   285
      Width           =   960
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   75
      Top             =   75
      Width           =   825
   End
   Begin VB.Image bg1 
      Height          =   690
      Left            =   1095
      Picture         =   "b8ToolButton.ctx":1466
      Stretch         =   -1  'True
      Top             =   1380
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image bg2 
      Height          =   690
      Left            =   2745
      Picture         =   "b8ToolButton.ctx":16D0
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image bg3 
      Height          =   690
      Left            =   1425
      Picture         =   "b8ToolButton.ctx":193A
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   0
      Picture         =   "b8ToolButton.ctx":1A34
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19995
   End
End
Attribute VB_Name = "b8ToolButton"
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
Event Declarations():
Event Click()

Dim m_ForeColor As OLE_COLOR

Private Sub bg1_Click()
    Ctl_OnDown
End Sub

Private Function Ctl_OnDown()
    Dim s As Single
    
    bg1.Picture = i1(1).Picture
    bg3.Picture = i2(1).Picture
    bg2.Picture = i3(1).Picture
    
    DoEvents
    
    
    
    s = GetTickCount + 40
    While GetTickCount < s
    Wend
    bg1.Picture = i1(0).Picture
    bg3.Picture = i2(0).Picture
    bg2.Picture = i3(0).Picture
    
    RaiseEvent Click
    
    
    
    
    
End Function

Private Sub bg2_Click()
    Ctl_OnDown
End Sub

Private Sub bg3_Click()
    Ctl_OnDown
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CtlMouseOver
End Sub

Private Sub imgIcon_Click()
    Ctl_OnDown
End Sub

Private Sub imgIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CtlMouseOver
End Sub

Private Sub lblCaption_Click()
    Ctl_OnDown
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CtlMouseOver
End Sub

Private Sub timerMouse_Timer()
    Dim p As POINTAPI
    Dim R As RECT

    GetWindowRect UserControl.hwnd, R
    GetCursorPos p
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
        timerMouse.Enabled = False
            bg1.Visible = False
            bg2.Visible = False
            bg3.Visible = False
            Image2.Visible = True
            
            UserControl.Parent.MousePointer = vbDefault
    End If
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CtlMouseOver
End Sub

Private Sub CtlMouseOver()
    bg1.Visible = True
    bg2.Visible = True
    bg3.Visible = True
    
    UserControl.Parent.MouseIcon = imgHand.Picture
    UserControl.Parent.MousePointer = vbCustom
    
    Image2.Visible = False
    timerMouse.Enabled = True
End Sub
Private Sub UserControl_Resize()
    Image2.Height = GetHeight
    bg1.Move 0, 0, bg1.Width, GetHeight
    bg2.Move GetWidth - bg2.Width, 0, bg2.Width, GetHeight
    bg3.Move bg1.Width, 0, GetWidth - bg2.Width - bg1.Width, GetHeight
    
    If imgIcon.Left + imgIcon.Width + 3 < GetWidth Then
        lblCaption.Move imgIcon.Left + imgIcon.Width + 3, (GetHeight / 2) - (lblCaption.Height / 2)
    Else
        lblCaption.Move (GetWidth - lblCaption.Height) / 2, (GetHeight / 2) - (lblCaption.Height / 2)
    End If
    
    'If imgDown.Visible = True Then
    '    imgDown.Move 0, 0, GetWidth, GetHeight
    'End If
End Sub



Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgIcon,imgIcon,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = imgIcon.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set imgIcon.Picture = New_Picture
    UserControl_Resize
    PropertyChanged "Picture"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set imgIcon.Picture = PropBag.ReadProperty("Picture", Nothing)
    lblCaption.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "b8ToolButton")
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontName = PropBag.ReadProperty("FontName", lblCaption.FontName)
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", lblCaption.FontSize)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_ForeColor = lblCaption.ForeColor
    lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set imgIcon2.Picture = PropBag.ReadProperty("DisabledPicture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", imgIcon.Picture, Nothing)
    Call PropBag.WriteProperty("Alignment", lblCaption.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "b8ToolButton")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("DisabledPicture", imgIcon2.Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Alignment
Public Property Get Alignment() As Integer
    Alignment = lblCaption.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    lblCaption.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = lblCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblCaption.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontName
Public Property Get FontName() As String
    FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblCaption.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = lblCaption.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblCaption.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = lblCaption.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblCaption.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    If New_Enabled = True Then
        lblCaption.ForeColor = m_ForeColor
        imgIcon.Visible = True
        imgIcon2.Visible = False
    Else
        lblCaption.ForeColor = &HBBCACD
        imgIcon.Visible = False
        imgIcon2.Move imgIcon.Top, imgIcon.Left
        imgIcon2.Visible = True
    End If
    
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgIcon2,imgIcon2,-1,Picture
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set DisabledPicture = imgIcon2.Picture
End Property

Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
    Set imgIcon2.Picture = New_DisabledPicture
    PropertyChanged "DisabledPicture"
End Property

