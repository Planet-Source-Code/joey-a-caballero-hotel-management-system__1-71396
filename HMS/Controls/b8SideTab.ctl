VERSION 5.00
Begin VB.UserControl b8SideTab 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ControlContainer=   -1  'True
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3465
      Top             =   2235
   End
   Begin VB.PictureBox bgCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H00C25418&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   660
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   1170
      Width           =   2955
      Begin VB.Image imgLeft 
         Height          =   345
         Left            =   2625
         Picture         =   "b8SideTab.ctx":0000
         Top             =   0
         Width           =   15
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "b8SideTab"
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
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   2685
      End
      Begin VB.Image imgTitleBG 
         Height          =   345
         Left            =   615
         Picture         =   "b8SideTab.ctx":0041
         Stretch         =   -1  'True
         Top             =   30
         Width           =   1290
      End
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   0
      Picture         =   "b8SideTab.ctx":00DF
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLeft1 
      Height          =   345
      Left            =   3375
      Picture         =   "b8SideTab.ctx":09A9
      Top             =   3195
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgLeft2 
      Height          =   345
      Left            =   2250
      Picture         =   "b8SideTab.ctx":09EA
      Top             =   3855
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgConLeft 
      Height          =   345
      Left            =   3780
      Picture         =   "b8SideTab.ctx":0A2B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   15
   End
   Begin VB.Image iLeft 
      Height          =   345
      Left            =   0
      Picture         =   "b8SideTab.ctx":0A6C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image imgbg3 
      Height          =   345
      Left            =   2700
      Picture         =   "b8SideTab.ctx":0AAD
      Stretch         =   -1  'True
      Top             =   4785
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgbg1 
      Height          =   345
      Left            =   2940
      Picture         =   "b8SideTab.ctx":0B4B
      Stretch         =   -1  'True
      Top             =   4185
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Image imgbg2 
      Height          =   345
      Left            =   885
      Picture         =   "b8SideTab.ctx":0BE9
      Stretch         =   -1  'True
      Top             =   4440
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Shape shpBorder 
      Height          =   1485
      Left            =   300
      Top             =   2460
      Width           =   1755
   End
End
Attribute VB_Name = "b8SideTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'code by:
'Vincent J. Jamero
'bob8choi@yahoo.com


Option Explicit
'Default Property Values:
Const m_def_ContractedForeColor = &HC25418
Const m_def_ExpandedForeColor = &HFFFFFF
Const m_def_Enabled = True
Const m_def_AutoExpand = True
Const m_def_ResizeAni = True
Const m_def_Expanded = False
Const m_def_MaxHeight = 0
'Property Variables:
Dim m_ContractedForeColor As OLE_COLOR
Dim m_ExpandedForeColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_AutoExpand As Boolean
Dim m_ResizeAni As Boolean
Dim m_Expanded As Boolean
Dim m_MaxHeight As Integer

'apis
Private Declare Function GetTickCount Lib "kernel32" () As Long


'Event Declarations:
Event CaptionMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblCaption,lblCaption,-1,MouseUp
Attribute CaptionMouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event CompleteContract()
Event CompleteExpand()
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event CaptionClick() 'MappingInfo=lblCaption,lblCaption,-1,Click
Attribute CaptionClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."




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




Private Sub imgTitle1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTitle1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseUp Button, Shift, X, Y
End Sub

Private Sub imgTitle2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgTitle2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseUp Button, Shift, X, Y
End Sub





Private Sub imgLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CaptionClick
End Sub

Private Sub imgLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MeMouseOnOver
End Sub

Private Sub imgLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseUp Button, Shift, X, Y
End Sub

Private Sub imgTitleBG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent CaptionClick
End Sub

Private Sub imgTitleBG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MeMouseOnOver
End Sub

Private Sub imgTitleBG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCaption_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    
    RaiseEvent CaptionClick
End Sub



Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MeMouseOnOver
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    Set UserControl.Font = New_Font
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MaxHeight() As Integer
    MaxHeight = m_MaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Integer)
    m_MaxHeight = New_MaxHeight
    PropertyChanged "MaxHeight"
End Property




'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MaxHeight = m_def_MaxHeight
    m_Expanded = m_def_Expanded
    m_ResizeAni = m_def_ResizeAni
    m_AutoExpand = m_def_AutoExpand
    m_Enabled = m_def_Enabled
    m_ContractedForeColor = m_def_ContractedForeColor
    m_ExpandedForeColor = m_def_ExpandedForeColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblCaption.Caption = PropBag.ReadProperty("Caption", "b8SideTab")
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblCaption.FontName = PropBag.ReadProperty("FontName", lblCaption.FontName)
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", lblCaption.FontSize)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H30A0B8)
    m_MaxHeight = PropBag.ReadProperty("MaxHeight", m_def_MaxHeight)
    m_Expanded = PropBag.ReadProperty("Expanded", m_def_Expanded)
    m_ResizeAni = PropBag.ReadProperty("ResizeAni", m_def_ResizeAni)
    shpBorder.BorderColor = PropBag.ReadProperty("BorderColor", &H80000008)
    m_AutoExpand = PropBag.ReadProperty("AutoExpand", m_def_AutoExpand)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_ContractedForeColor = PropBag.ReadProperty("ContractedForeColor", m_def_ContractedForeColor)
    m_ExpandedForeColor = PropBag.ReadProperty("ExpandedForeColor", m_def_ExpandedForeColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
End Sub

Private Sub UserControl_Resize()
    
On Error Resume Next

    iLeft.Move 0, 0, iLeft.Width, GetHeight

    bgCaption.Move iLeft.Width, 0, GetWidth - iLeft.Width
    
    imgLeft.Move bgCaption.Width - imgLeft.Width, 0
    
    imgTitleBG.Move 0, 0, bgCaption.Width - imgLeft.Width
    
    lblCaption.Move iLeft.Width, 4, GetWidth

    imgConLeft.Move GetWidth - imgConLeft.Width, 0, imgConLeft.Width, GetHeight

        
    shpBorder.Move iLeft.Width, 0, GetWidth, GetHeight
     
    
    RaiseEvent Resize
    

End Sub





'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "b8SideTab")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H30A0B8)
    Call PropBag.WriteProperty("MaxHeight", m_MaxHeight, m_def_MaxHeight)
    Call PropBag.WriteProperty("Expanded", m_Expanded, m_def_Expanded)
    Call PropBag.WriteProperty("ResizeAni", m_ResizeAni, m_def_ResizeAni)
    Call PropBag.WriteProperty("BorderColor", shpBorder.BorderColor, &H80000008)
    Call PropBag.WriteProperty("AutoExpand", m_AutoExpand, m_def_AutoExpand)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("ContractedForeColor", m_ContractedForeColor, m_def_ContractedForeColor)
    Call PropBag.WriteProperty("ExpandedForeColor", m_ExpandedForeColor, m_def_ExpandedForeColor)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
End Sub

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelY
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelX
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,false
Public Property Get Expanded() As Boolean
Attribute Expanded.VB_MemberFlags = "400"
    Expanded = m_Expanded
End Property

Public Property Let Expanded(ByVal New_Expanded As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    
    Dim NewHeight As Integer
    Dim st As Single
    Dim StepSize As Integer
    Dim oldColor As OLE_COLOR
    Dim ContractSize As Integer
   
    If New_Expanded = False Then
    
        UserControl.Height = Screen.TwipsPerPixelY * (bgCaption.Height)
        m_Expanded = False
        Set imgTitleBG.Picture = imgbg2.Picture
        Set imgLeft.Picture = imgLeft2.Picture
        lblCaption.ForeColor = ContractedForeColor
        RaiseEvent CompleteContract
    Else
    
        'set images
        Set imgTitleBG.Picture = imgbg1.Picture
        Set imgLeft.Picture = imgLeft1.Picture
        imgConLeft.Move GetWidth - imgConLeft.Width, 0, imgConLeft.Width, GetHeight
        
        If ResizeAni = True Then
            
            NewHeight = MaxHeight
            
            
            If NewHeight > UserControl.Height Then
            
                
                StepSize = (NewHeight - UserControl.Height) / Screen.TwipsPerPixelY * 2
            
                While UserControl.Height < NewHeight
                
                    UserControl.Height = UserControl.Height + StepSize
                    DoEvents
                    
                    st = GetTickCount + 4
                    While st > GetTickCount
                        
                    Wend
                Wend


                m_Expanded = True
                Set imgTitleBG.Picture = imgbg1.Picture
                Set imgLeft.Picture = imgLeft1.Picture
                imgConLeft.Move GetWidth - imgConLeft.Width, 0, imgConLeft.Width, GetHeight
                lblCaption.ForeColor = ExpandedForeColor
                RaiseEvent CompleteExpand
                
            Else
            
               m_Expanded = False
            End If
            
        Else
            UserControl.Height = MaxHeight
        End If
        
    End If
    
    
    
    PropertyChanged "Expanded"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get ResizeAni() As Boolean
    ResizeAni = m_ResizeAni
End Property

Public Property Let ResizeAni(ByVal New_ResizeAni As Boolean)
    m_ResizeAni = New_ResizeAni
    PropertyChanged "ResizeAni"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BorderColor() As OLE_COLOR
'    BorderColor = UserControl.BackColor
'End Property
'
'Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
'    UserControl.BackColor() = New_BorderColor
'    PropertyChanged "BorderColor"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBorder,shpBorder,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shpBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get AutoExpand() As Boolean
    AutoExpand = m_AutoExpand
End Property

Public Property Let AutoExpand(ByVal New_AutoExpand As Boolean)
    m_AutoExpand = New_AutoExpand
    PropertyChanged "AutoExpand"
End Property

















Public Sub HideExpand()
    
    Dim NewHeight As Integer
    Dim st As Single
    Dim StepSize As Integer
    Dim oldColor As OLE_COLOR

    
    If Expanded = True Then
        UserControl.Height = Screen.TwipsPerPixelX * (bgCaption.Height)
        m_Expanded = False
        Set imgTitleBG.Picture = imgbg2.Picture
        Set imgLeft.Picture = imgLeft2.Picture
        lblCaption.ForeColor = ContractedForeColor
        RaiseEvent CompleteContract
    Else
        
        If ResizeAni = True Then
            NewHeight = MaxHeight
            If NewHeight > UserControl.Height Then
            

                
                StepSize = (NewHeight - UserControl.Height) / Screen.TwipsPerPixelY * 2
                While UserControl.Height < NewHeight
                
                    UserControl.Height = UserControl.Height + StepSize
                    st = GetTickCount + 4
                    While st > GetTickCount
                        DoEvents
                    Wend
                Wend

    
                m_Expanded = True
                Set imgTitleBG.Picture = imgbg1.Picture
                Set imgLeft.Picture = imgLeft1.Picture
                imgConLeft.Move GetWidth - imgConLeft.Width, 0, imgConLeft.Width, GetHeight
                lblCaption.ForeColor = ExpandedForeColor
                RaiseEvent CompleteExpand
            Else
                m_Expanded = False
                lblCaption.ForeColor = ContractedForeColor
            End If
            
        Else
            UserControl.Height = MaxHeight
            m_Expanded = True
            lblCaption.ForeColor = ExpandedForeColor
            RaiseEvent CompleteExpand
        End If
        
    End If

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Dim ConCtrl As Control
    On Error Resume Next
    
    For Each ConCtrl In UserControl.ContainedControls
        ConCtrl.Enabled = New_Enabled
    Next
    
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get ContainedControls() As Object
Attribute ContainedControls.VB_Description = "A collection whose elements represent each control on a form, including elements of control arrays. "
    Set Controls = UserControl.ContainedControls
    
End Property




Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_AutoExpand = True Then
        If m_Expanded = True Then
            Expanded = False
        Else
            Expanded = True
        End If
    End If
    RaiseEvent CaptionMouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ContractedForeColor() As OLE_COLOR
    ContractedForeColor = m_ContractedForeColor
End Property

Public Property Let ContractedForeColor(ByVal New_ContractedForeColor As OLE_COLOR)
    m_ContractedForeColor = New_ContractedForeColor
    PropertyChanged "ContractedForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ExpandedForeColor() As OLE_COLOR
    ExpandedForeColor = m_ExpandedForeColor
End Property

Public Property Let ExpandedForeColor(ByVal New_ExpandedForeColor As OLE_COLOR)
    m_ExpandedForeColor = New_ExpandedForeColor
    PropertyChanged "ExpandedForeColor"
End Property


Private Function MeMouseOnOver()
    UserControl.Parent.MouseIcon = imgHand.Picture
    UserControl.Parent.MousePointer = vbCustom

    imgTitleBG.Picture = imgbg3.Picture
    timerMouse.Enabled = True
End Function
Private Sub timerMouse_Timer()
    Dim p As POINTAPI
    Dim R As RECT

    GetWindowRect bgCaption.hwnd, R
    GetCursorPos p
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
        timerMouse.Enabled = False
        
        UserControl.Parent.MousePointer = vbDefault
        
            If Expanded = True Then
                imgTitleBG.Picture = imgbg1.Picture
            Else
                imgTitleBG.Picture = imgbg2.Picture
            End If
            
    End If
End Sub
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

