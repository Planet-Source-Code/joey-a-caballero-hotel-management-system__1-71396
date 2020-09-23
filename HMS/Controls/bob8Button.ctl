VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl bob8Button 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   Begin MSComctlLib.ImageList imglistButtonColor 
      Left            =   3465
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   34
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":0000
            Key             =   "DarkGray"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":01B2
            Key             =   "LihtGray"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":0363
            Key             =   "DarkBlue"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":052D
            Key             =   "LightBlue"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":06E0
            Key             =   "DarkGreen"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":08A8
            Key             =   "LightGreen"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":0A5A
            Key             =   "DarkRed"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":0C28
            Key             =   "LightRed"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":0DF1
            Key             =   "DarkYellow"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bob8Button.ctx":0FC3
            Key             =   "LightYellow"
         EndProperty
      EndProperty
   End
   Begin VB.Shape imgSpacer2 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   1785
      Top             =   0
      Width           =   15
   End
   Begin VB.Shape imgSpacer1 
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   555
      Top             =   0
      Width           =   15
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   945
      TabIndex        =   0
      Top             =   135
      Width           =   450
   End
   Begin VB.Image imgSpacerRight 
      Height          =   510
      Left            =   1800
      Picture         =   "bob8Button.ctx":1179
      Stretch         =   -1  'True
      Top             =   0
      Width           =   540
   End
   Begin VB.Image imgSpacerLeft 
      Height          =   510
      Left            =   0
      Picture         =   "bob8Button.ctx":131A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   540
   End
   Begin VB.Image imgButton 
      Height          =   510
      Left            =   540
      Picture         =   "bob8Button.ctx":14BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "bob8Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum eColorScheme
    Default = 0
    Gray = 1
    Blue = 2
    Green = 3
    Red = 4
    Yelow = 5
End Enum
Private Const MaxColorScheme = 5


'Default Property Values:
Const m_def_ColorScheme = 0
Const m_def_ButtonWidth = 82
Const m_def_ButtonAlignment = 2
'Const m_def_ColorScheme = 0
'Property Variables:
Dim m_ColorScheme As Integer
Dim m_ButtonWidth As Variant
Dim m_ButtonAlignment As Integer
'Dim m_ColorScheme As Variant

















Private Sub AlignCaption()
    lblCaption.Left = imgButton.Left
    lblCaption.Width = imgButton.Width
    lblCaption.Top = imgButton.Top + (imgButton.Height / 2) - (lblCaption.Height / 2)
    PropertyChanged "Font"
End Sub

Private Sub AlignOnResize()
    'MsgBox UserControl.Width / UserControl.ScaleWidth
   'align images
    If UserControl.ScaleWidth < 1 Then '0 or auto size depeds on the size of button
        UserControl.Width = imgButton.Width * 15
         'UserControl.ScaleWidth * Screen.TwipsPerPixelX
    Else
        'UserControl.Width = 100
    End If
    
    
    
On Error Resume Next
    'set top
    imgSpacerLeft.Top = 0
    imgSpacer1.Top = 0
    imgButton.Top = 0
    imgSpacer2.Top = 0
    imgSpacerRight.Top = 0
    'resize height
    imgSpacerLeft.Height = UserControl.ScaleHeight
    imgSpacer1.Height = UserControl.ScaleHeight
    imgButton.Height = UserControl.ScaleHeight
    imgSpacer2.Height = UserControl.ScaleHeight
    imgSpacerRight.Height = UserControl.ScaleHeight

        Select Case ButtonAlignment
            Case 0 'left
            
            Case 1 'right
            
            Case 2 'center

                'set left spacer
                imgSpacerLeft.Left = 0
                imgSpacerLeft.Width = (UserControl.ScaleWidth - (imgSpacer1.Width + imgButton.Width + imgSpacer2.Width)) / 2
                Debug.Print "US" & UserControl.ScaleWidth
                Debug.Print " - " & (imgSpacer1.Width + imgButton.Width + imgSpacer2.Width)
                Debug.Print "IS" & imgSpacerLeft.Width
                'set spacer1
                    imgSpacer1.Left = imgSpacerLeft.Width
                'set the button
                    imgButton.Left = imgSpacer1.Left + 1

                'setspacer2
                    imgSpacer2.Left = imgButton.Left + imgButton.Width
                'set right spacer
                imgSpacerRight.Left = imgSpacer2.Left + 1
                imgSpacerRight.Width = (UserControl.ScaleWidth - (imgSpacer1.Width + imgButton.Width + imgSpacer2.Width)) / 2
                
            Case Else 'error
                'raise error here
        End Select
        Call AlignCaption
End Sub

Private Sub ChangeColor(NewColor As Integer)
    imgSpacerLeft.Picture = imglistButtonColor.ListImages(2).Picture
    imgButton.Picture = imglistButtonColor.ListImages(NewColor * 2).Picture
    imgSpacerRight.Picture = imglistButtonColor.ListImages(2).Picture
End Sub




Private Sub ButtonOnClick()

End Sub













'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    'align caption
    Call AlignCaption
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblCaption.FontName() = New_FontName
    'align caption
    Call AlignCaption
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = lblCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblCaption.FontItalic() = New_FontItalic
    'align caption
    Call AlignCaption
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    'align caption
    Call AlignCaption
    PropertyChanged "FontBold"
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
    
    'align caption
    Call AlignCaption
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get ColorScheme() As Variant
'    ColorScheme = m_ColorScheme
'End Property
'
'Public Property Let ColorScheme(ByVal New_ColorScheme As Variant)
'    m_ColorScheme = New_ColorScheme
'    PropertyChanged "ColorScheme"
'End Property






'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_ColorScheme = m_def_ColorScheme
    m_ButtonAlignment = m_def_ButtonAlignment
    m_ColorScheme = m_def_ColorScheme
    m_ButtonWidth = m_def_ButtonWidth
    m_BackWidth = m_def_BackWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.FontSize = PropBag.ReadProperty("FontSize", 10)
    lblCaption.FontName = PropBag.ReadProperty("FontName", "Tahoma")
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_ColorScheme = PropBag.ReadProperty("ColorScheme", m_def_ColorScheme)
    m_ButtonAlignment = PropBag.ReadProperty("ButtonAlignment", m_def_ButtonAlignment)
    m_ColorScheme = PropBag.ReadProperty("ColorScheme", m_def_ColorScheme)
    m_ButtonWidth = PropBag.ReadProperty("ButtonWidth", m_def_ButtonWidth)
    m_BackWidth = PropBag.ReadProperty("BackWidth", m_def_BackWidth)
End Sub



Private Sub UserControl_Resize()
    Call AlignOnResize
End Sub

Private Sub UserControl_Show()
    'refresh
    Call AlignOnResize
    Call ChangeColor(ColorScheme)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
'    Call PropBag.WriteProperty("ColorScheme", m_ColorScheme, m_def_ColorScheme)
    Call PropBag.WriteProperty("ButtonAlignment", m_ButtonAlignment, m_def_ButtonAlignment)
    Call PropBag.WriteProperty("ColorScheme", m_ColorScheme, m_def_ColorScheme)
    Call PropBag.WriteProperty("ButtonWidth", m_ButtonWidth, m_def_ButtonWidth)
    Call PropBag.WriteProperty("BackWidth", m_BackWidth, m_def_BackWidth)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ButtonAlignment() As Integer
    ButtonAlignment = m_ButtonAlignment
End Property

Public Property Let ButtonAlignment(ByVal New_ButtonAlignment As Integer)
    m_ButtonAlignment = New_ButtonAlignment
    PropertyChanged "ButtonAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ColorScheme() As Integer
    ColorScheme = m_ColorScheme
End Property

Public Property Let ColorScheme(ByVal New_ColorScheme As Integer)
    If New_ColorScheme > MaxColorScheme Or New_ColorScheme < 0 Then
        New_ColorScheme = 0
        Exit Property
    End If
    
    m_ColorScheme = New_ColorScheme
    Call ChangeColor(New_ColorScheme)
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,82
Public Property Get ButtonWidth() As Variant
    ButtonWidth = m_ButtonWidth
End Property

Public Property Let ButtonWidth(ByVal New_ButtonWidth As Variant)
    If New_ButtonWidth < 1 Then
        'raise error here
        Exit Property
    End If
    
    imgButton.Width = New_ButtonWidth
    Call AlignOnResize
    
    m_ButtonWidth = New_ButtonWidth
    PropertyChanged "ButtonWidth"
End Property



