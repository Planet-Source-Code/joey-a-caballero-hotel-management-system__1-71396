VERSION 5.00
Begin VB.UserControl b8ChildTitleBar 
   BackColor       =   &H00C25418&
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   Begin VB.PictureBox imgClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   2490
      Picture         =   "bgChildTitleBar.ctx":0000
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      ToolTipText     =   "Close This Window"
      Top             =   660
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4005
      Top             =   660
   End
   Begin VB.PictureBox imgClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3180
      Picture         =   "bgChildTitleBar.ctx":046A
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   2
      ToolTipText     =   "Close This Window"
      Top             =   705
      Width           =   270
   End
   Begin VB.Image imgGradTheme 
      Height          =   345
      Index           =   2
      Left            =   480
      Picture         =   "bgChildTitleBar.ctx":07F6
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   705
   End
   Begin VB.Image imgGradTheme 
      Height          =   285
      Index           =   0
      Left            =   3690
      Top             =   1290
      Width           =   1470
   End
   Begin VB.Image imgGradTheme 
      Height          =   345
      Index           =   1
      Left            =   1755
      Picture         =   "bgChildTitleBar.ctx":0894
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "b8Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   585
   End
   Begin VB.Image imgGrad 
      Height          =   435
      Left            =   1005
      Picture         =   "bgChildTitleBar.ctx":0936
      Stretch         =   -1  'True
      Top             =   15
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "b8ChildTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_CloseButton = True
Const m_def_GradTheme = 0
'Property Variables:
Dim m_CloseButton As Boolean
Dim m_GradTheme As Integer

Public Enum eGradTheme
    No_Theme = 0
    MCE_Blue = 1
    Gray = 2
End Enum

Dim MouseOnDOwn As Boolean

Private Sub imgClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(1).Visible = False
    'RaiseEvent CloseMouseDown(Button, Shift, X, Y)
    
    MouseOnDOwn = True
End Sub

Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(1).Visible = IIf(MouseOnDOwn, False, True)
    timerMouse.Enabled = True
End Sub

Private Sub imgClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim p As POINTAPI
    Dim R As RECT
    
    imgClose(1).Visible = True
    
    GetWindowRect imgClose(0).hwnd, R
    GetCursorPos p
    
    If Not (p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom) Then
        
        'RaiseEvent CLoseClick
        
            On Error Resume Next
            Unload UserControl.Parent
    End If
    
    
    'RaiseEvent CloseMouseUp(Button, Shift, X, Y)
    MouseOnDOwn = False
    
End Sub



Private Sub timerMouse_Timer()
    Dim p As POINTAPI
    Dim R As RECT

    GetWindowRect imgClose(0).hwnd, R
    GetCursorPos p
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
        timerMouse.Enabled = False
        imgClose(1).Visible = False
    End If
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
    
    imgGrad.Move 0, 0, GetWidth, GetHeight
    
    lblCaption.Top = (GetHeight / 2) - (lblCaption.Height / 2)
    
    imgClose(0).Move GetWidth - 2 - imgClose(0).Width, 2
    imgClose(1).Move GetWidth - 2 - imgClose(1).Width, 2
End Sub
Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = lblCaption.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    lblCaption.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = lblCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblCaption.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblCaption.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = lblCaption.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblCaption.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = lblCaption.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblCaption.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Title")
    Set lblCaption.Font = PropBag.ReadProperty("Font", lblCaption.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontName = PropBag.ReadProperty("FontName", lblCaption.FontName)
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", lblCaption.FontSize)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_GradTheme = PropBag.ReadProperty("GradTheme", m_def_GradTheme)
    m_CloseButton = PropBag.ReadProperty("CloseButton", m_def_CloseButton)
End Sub

Private Sub UserControl_Show()
    
    Select Case m_GradTheme
        Case No_Theme
            imgGrad.Visible = False
        
        Case Else
            If CloseButton = True Then
            
                imgGrad.Picture = imgGradTheme(m_GradTheme).Picture
                imgGrad.Move 0, 0, GetWidth, GetHeight
                imgGrad.Visible = True
                
            Else
            
                imgClose(0).Visible = False
                imgClose(1).Visible = False
            End If
            
    End Select
    
    If CloseButton = False Then

            
                imgClose(0).Visible = False
                imgClose(1).Visible = False
            End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", lblCaption.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Label1")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("GradTheme", m_GradTheme, m_def_GradTheme)
    Call PropBag.WriteProperty("CloseButton", m_CloseButton, m_def_CloseButton)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get GradTheme() As eGradTheme
    GradTheme = m_GradTheme
End Property

Public Property Let GradTheme(ByVal New_GradTheme As eGradTheme)
    m_GradTheme = New_GradTheme
    
    Select Case New_GradTheme
        Case No_Theme
            imgGrad.Visible = False
        
        Case Else
            imgGrad.Picture = imgGradTheme(New_GradTheme).Picture
            imgGrad.Move 1, 0, GetWidth - 1, GetHeight
            imgGrad.Visible = True

        
    End Select
    
    PropertyChanged "GradTheme"
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_GradTheme = m_def_GradTheme
    m_CloseButton = m_def_CloseButton
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get CloseButton() As Boolean
    CloseButton = m_CloseButton
End Property

Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
    m_CloseButton = New_CloseButton
    imgClose(0).Visible = New_CloseButton
    imgClose(1).Visible = New_CloseButton
    PropertyChanged "CloseButton"
End Property

