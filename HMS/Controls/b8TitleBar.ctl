VERSION 5.00
Begin VB.UserControl b8TitleBar 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   Begin VB.PictureBox imgClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   2970
      Picture         =   "b8TitleBar.ctx":0000
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      ToolTipText     =   "Close This Window"
      Top             =   30
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox imgClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   3660
      Picture         =   "b8TitleBar.ctx":0396
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   0
      ToolTipText     =   "Close This Window"
      Top             =   60
      Width           =   270
   End
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4485
      Top             =   15
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "b8TitleBar"
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
      Left            =   375
      TabIndex        =   2
      Top             =   75
      Width           =   870
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   45
      Picture         =   "b8TitleBar.ctx":0722
      Stretch         =   -1  'True
      Top             =   60
      Width           =   240
   End
   Begin VB.Image imgShadow 
      Height          =   345
      Index           =   1
      Left            =   1755
      Picture         =   "b8TitleBar.ctx":0CAC
      Top             =   2790
      Width           =   315
   End
   Begin VB.Image imgBGRight 
      Height          =   345
      Index           =   1
      Left            =   3570
      Picture         =   "b8TitleBar.ctx":0F8E
      Top             =   2355
      Width           =   30
   End
   Begin VB.Image imgBG 
      Height          =   345
      Index           =   1
      Left            =   1470
      Picture         =   "b8TitleBar.ctx":1141
      Stretch         =   -1  'True
      Top             =   2250
      Width           =   1590
   End
   Begin VB.Image imgBGLeft 
      Height          =   345
      Index           =   1
      Left            =   1095
      Picture         =   "b8TitleBar.ctx":12CB
      Top             =   2175
      Width           =   30
   End
   Begin VB.Image imgBGRight 
      Height          =   345
      Index           =   0
      Left            =   0
      Picture         =   "b8TitleBar.ctx":148D
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgBGLeft 
      Height          =   345
      Index           =   0
      Left            =   0
      Picture         =   "b8TitleBar.ctx":164A
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgShadow 
      Height          =   345
      Index           =   0
      Left            =   15
      Picture         =   "b8TitleBar.ctx":183F
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgBG 
      Height          =   345
      Index           =   0
      Left            =   15
      Picture         =   "b8TitleBar.ctx":1B46
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5310
   End
End
Attribute VB_Name = "b8TitleBar"
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

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long




'events
Public Event CloseMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event CloseMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event CLoseClick()



Dim MouseOnDOwn As Boolean
Dim iCurIMGIndex As Integer


'Default Property Values:
Const m_def_AutoFunction = True
'Property Variables:
Dim m_ShadowVisible As Boolean
Dim m_AutoFunction As Boolean
'Event Declarations:
Event DblClick() 'MappingInfo=imgBG,imgBG,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."


Private Sub FormDrag(frmName As Form) 'procedure to drag a no-titlebar form
    ReleaseCapture
    Call SendMessage(frmName.hwnd, &HA1, 2, 0&)
End Sub


Private Sub imgBG_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(1).Visible = False
    RaiseEvent CloseMouseDown(Button, Shift, X, Y)
    
    MouseOnDOwn = True
End Sub

Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(1).Visible = IIf(MouseOnDOwn, False, True)
    timerMouse.Enabled = True
End Sub

Private Sub imgClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim P As POINTAPI
    Dim R As RECT
    
    imgClose(1).Visible = True
    
    GetWindowRect imgClose(0).hwnd, R
    GetCursorPos P
    
    If Not (P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom) Then
        
        RaiseEvent CLoseClick
        
        If AutoFunction = True Then
            On Error Resume Next
            Unload UserControl.Parent
        End If
    End If
    
    
    RaiseEvent CloseMouseUp(Button, Shift, X, Y)
    MouseOnDOwn = False
    
End Sub



Private Sub timerMouse_Timer()
    Dim P As POINTAPI
    Dim R As RECT

    GetWindowRect imgClose(0).hwnd, R
    GetCursorPos P
    
    If P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom Then
        timerMouse.Enabled = False
        imgClose(1).Visible = False
    End If
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_Initialize()
    iCurIMGIndex = 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And AutoFunction = True Then
        FormDrag UserControl.Parent
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    UserControl.Height = 23 * Screen.TwipsPerPixelY
    
    imgBGLeft(iCurIMGIndex).Move 0, 0
    imgBGLeft(iCurIMGIndex).Visible = True
    
    imgBG(iCurIMGIndex).Move 2, 0, GetWidth - 2, GetHeight
    imgBG(iCurIMGIndex).Visible = True
    
    imgBGRight(iCurIMGIndex).Move GetWidth - 2, 0
    imgBGRight(iCurIMGIndex).Visible = True
    
    imgIcon.Move 5, 3
    
    If m_ShadowVisible = True Then
        imgShadow(iCurIMGIndex).Move 1, 0
        imgShadow(iCurIMGIndex).Visible = True
    End If
    
    imgClose(0).Move GetWidth - 2 - imgClose(0).Width, 2
    imgClose(1).Move GetWidth - 2 - imgClose(1).Width, 2

End Sub


Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelY
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelX
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    
    If New_Caption = lblCaption.Caption Then Exit Property
    
    lblCaption.Caption() = New_Caption
    
    If AutoFunction = True Then
        UserControl.Parent.Caption = New_Caption
    End If
    
    PropertyChanged "Caption"
End Property

Private Sub imgBG_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
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
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoFunction() As Boolean
    AutoFunction = m_AutoFunction
End Property

Public Property Let AutoFunction(ByVal New_AutoFunction As Boolean)
    m_AutoFunction = New_AutoFunction
    PropertyChanged "AutoFunction"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AutoFunction = m_def_AutoFunction
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.Caption = PropBag.ReadProperty("Caption", UserControl.Parent.Caption)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontName = PropBag.ReadProperty("FontName", lblCaption.FontName)
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", lblCaption.FontSize)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    m_AutoFunction = PropBag.ReadProperty("AutoFunction", m_def_AutoFunction)
    Set imgIcon.Picture = PropBag.ReadProperty("Icon", Nothing)
    imgShadow(iCurIMGIndex).Visible = PropBag.ReadProperty("ShadowVisible", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", lblCaption.Caption, UserControl.Parent.Caption)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("AutoFunction", m_AutoFunction, m_def_AutoFunction)
    Call PropBag.WriteProperty("Icon", imgIcon.Picture)
    Call PropBag.WriteProperty("ShadowVisible", imgShadow(iCurIMGIndex).Visible, True)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgIcon,imgIcon,-1,Picture
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Icon = imgIcon.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set imgIcon.Picture = New_Icon
    PropertyChanged "Icon"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get ShadowVisible() As Boolean
    ShadowVisible = imgShadow(iCurIMGIndex).Visible
End Property

Public Property Let ShadowVisible(ByVal New_ShadowVisible As Boolean)
    imgShadow(iCurIMGIndex).Visible = New_ShadowVisible
    m_ShadowVisible = New_ShadowVisible
    PropertyChanged "ShadowVisible"
End Property


Public Sub ParentFocus(Optional OnFocus As Boolean = True)
    If OnFocus = True Then
        iCurIMGIndex = 0
        imgBGLeft(1).Visible = False
        imgBG(1).Visible = False
        imgBGRight(1).Visible = False
        imgShadow(1).Visible = False
    Else
        iCurIMGIndex = 1
        imgBGLeft(0).Visible = False
        imgBG(0).Visible = False
        imgBGRight(0).Visible = False
        imgShadow(0).Visible = False
        
    End If
    
    Call UserControl_Resize
End Sub
