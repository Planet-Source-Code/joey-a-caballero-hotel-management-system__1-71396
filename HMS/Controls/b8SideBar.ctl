VERSION 5.00
Begin VB.UserControl b8SideBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00FAE5D3&
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FAE5D3&
   KeyPreview      =   -1  'True
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   239
   Begin VB.Timer timerMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2565
      Top             =   840
   End
   Begin VB.PictureBox bgDOWN 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3360
      Picture         =   "b8SideBar.ctx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox bgUP 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3330
      Picture         =   "b8SideBar.ctx":0284
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00F4DAC5&
      X1              =   208
      X2              =   208
      Y1              =   83
      Y2              =   337
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00EED0B8&
      X1              =   209
      X2              =   209
      Y1              =   72
      Y2              =   326
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E6C2A6&
      X1              =   210
      X2              =   210
      Y1              =   69
      Y2              =   323
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D16D2B&
      X1              =   211
      X2              =   211
      Y1              =   75
      Y2              =   329
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C25418&
      X1              =   212
      X2              =   212
      Y1              =   73
      Y2              =   327
   End
End
Attribute VB_Name = "b8SideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'code by: Vincent Jamero

Const ScrollStepSize = 15

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

'Event Declarations:
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event CtlsScroll()

Dim MouseDown As Boolean
Dim ButtonIndex As Integer






Public Function MoveUpControls(Optional iStep As Integer = 15)
    Dim ConCtrl As Control
    'check
    For Each ConCtrl In UserControl.ContainedControls
        ConCtrl.Top = ConCtrl.Top - iStep

    Next
    
    RaiseEvent CtlsScroll
End Function

Public Function MoveDownControls(Optional iStep As Integer = 15)
    Dim ConCtrl As Control
    'check
    For Each ConCtrl In UserControl.ContainedControls
        ConCtrl.Top = ConCtrl.Top + iStep

    Next
    
    RaiseEvent CtlsScroll
End Function






Private Function ScrollUp()
    Dim ConCtrl As Control
    Dim HighestSpace As Integer
    'check
    On Error Resume Next
    For Each ConCtrl In UserControl.ContainedControls
        If ConCtrl.Top < HighestSpace Then
            HighestSpace = ConCtrl.Top
        End If
    Next
    

    If HighestSpace < 0 Then
        If (HighestSpace + (Screen.TwipsPerPixelX * ScrollStepSize)) < 1 Then
            For Each ConCtrl In UserControl.ContainedControls
                ConCtrl.Top = ConCtrl.Top + (Screen.TwipsPerPixelX * ScrollStepSize)
            Next
        Else
            For Each ConCtrl In UserControl.ContainedControls
                ConCtrl.Top = ConCtrl.Top - HighestSpace
            Next
        End If
    End If
    
    CheckExceedControl
    
    RaiseEvent CtlsScroll
End Function
Private Function ScrollDown()
    Dim ConCtrl As Control
    Dim HighestSpace As Integer
    'check
    On Error Resume Next
    For Each ConCtrl In UserControl.ContainedControls
        If (ConCtrl.Top + ConCtrl.Height) - UserControl.Height > HighestSpace Then
            HighestSpace = (ConCtrl.Top + ConCtrl.Height) - UserControl.Height
        End If
    Next
    

    If HighestSpace > 0 Then
    
        If (HighestSpace - (Screen.TwipsPerPixelX * ScrollStepSize)) > 0 Then
            For Each ConCtrl In UserControl.ContainedControls
                ConCtrl.Top = ConCtrl.Top - (Screen.TwipsPerPixelX * ScrollStepSize)
            Next
        Else
            For Each ConCtrl In UserControl.ContainedControls
                ConCtrl.Top = ConCtrl.Top - HighestSpace
            Next
        End If
    End If
    
    CheckExceedControl
    
    RaiseEvent CtlsScroll
End Function
Public Sub CheckExceedControl()
    Dim ConCtrl As Control
    Dim FoundExceedTop As Boolean
    Dim FoundExceedBottom As Boolean
    On Error Resume Next
    For Each ConCtrl In UserControl.ContainedControls
    
         If ConCtrl.Visible = True And ConCtrl.Top < 0 Then
            bgUP.Move UserControl.ScaleWidth - bgUP.Width, 0
            bgUP.Visible = True
            bgUP.ZOrder 0
            FoundExceedTop = True
            Exit For
            
         End If
    Next
    
    For Each ConCtrl In UserControl.ContainedControls
         If ConCtrl.Visible = True And ConCtrl.Top + ConCtrl.Height > UserControl.Height Then
            bgDOWN.Move UserControl.ScaleWidth - bgDOWN.Width, UserControl.ScaleHeight - bgDOWN.Height
            bgDOWN.Visible = True
            bgDOWN.ZOrder 0
            FoundExceedBottom = True
            Exit For
            
         End If
    Next
    

    If FoundExceedTop = False Then
        bgUP.Visible = False
    End If
    If FoundExceedBottom = False Then
        bgDOWN.Visible = False
    End If
End Sub

Private Sub bgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ScrollUp
    MouseDown = True
    ButtonIndex = 1
End Sub

Private Sub bgDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonIndex = 1
    timerMouse.Enabled = True
End Sub

Private Sub bgDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonIndex = -1
    MouseDown = False
End Sub

Private Sub bgUP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ScrollUp
    MouseDown = True
    ButtonIndex = 0
End Sub

Private Sub bgUP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonIndex = 0
    timerMouse.Enabled = True
End Sub

Private Sub bgUP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonIndex = -1
    MouseDown = False
End Sub

Private Sub timerMouse_Timer()
    Dim p As POINTAPI
    Dim R As RECT

    If ButtonIndex = 0 Then
        GetWindowRect bgUP.hwnd, R
    Else
        GetWindowRect bgDOWN.hwnd, R
    End If
    
    GetCursorPos p
    
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
        timerMouse.Enabled = False
    Else
    
        If MouseDown = True Then
            If ButtonIndex = 0 Then
                ScrollUp
            Else
                ScrollDown
            End If
        End If
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            ScrollUp
        Case vbKeyDown
            ScrollDown
    End Select

End Sub



Private Sub UserControl_Paint()
    CheckExceedControl
End Sub

Private Sub UserControl_Resize()
    Line1.X1 = GetWidth - 1
    Line1.X2 = GetWidth - 1
    Line1.Y1 = 0
    Line1.Y2 = GetHeight
    
    Line2.X1 = GetWidth - 2
    Line2.X2 = GetWidth - 2
    Line2.Y1 = 0
    Line2.Y2 = GetHeight
    
    Line3.X1 = GetWidth - 3
    Line3.X2 = GetWidth - 3
    Line3.Y1 = 0
    Line3.Y2 = GetHeight
    
    Line4.X1 = GetWidth - 4
    Line4.X2 = GetWidth - 4
    Line4.Y1 = 0
    Line4.Y2 = GetHeight
    
    Line5.X1 = GetWidth - 5
    Line5.X2 = GetWidth - 5
    Line5.Y1 = 0
    Line5.Y2 = GetHeight

    CheckExceedControl
    RaiseEvent Resize
End Sub

Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000004)
   
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Line1.BorderColor = PropBag.ReadProperty("BorderColor1", 12735512)
    Line2.BorderColor = PropBag.ReadProperty("BorderColor2", 13724971)
    Line3.BorderColor = PropBag.ReadProperty("BorderColor3", 15123110)
    Line4.BorderColor = PropBag.ReadProperty("BorderColor4", 15651000)
    Line5.BorderColor = PropBag.ReadProperty("BorderColor5", 16046789)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000004)
   
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderColor1", Line1.BorderColor, 12735512)
    Call PropBag.WriteProperty("BorderColor2", Line2.BorderColor, 13724971)
    Call PropBag.WriteProperty("BorderColor3", Line3.BorderColor, 15123110)
    Call PropBag.WriteProperty("BorderColor4", Line4.BorderColor, 15651000)
    Call PropBag.WriteProperty("BorderColor5", Line5.BorderColor, 16046789)
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderColor
Public Property Get BorderColor1() As OLE_COLOR
Attribute BorderColor1.VB_Description = "Returns/sets the color of an object's border."
    BorderColor1 = Line1.BorderColor
End Property

Public Property Let BorderColor1(ByVal New_BorderColor1 As OLE_COLOR)
    Line1.BorderColor() = New_BorderColor1
    PropertyChanged "BorderColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line2,Line2,-1,BorderColor
Public Property Get BorderColor2() As OLE_COLOR
Attribute BorderColor2.VB_Description = "Returns/sets the color of an object's border."
    BorderColor2 = Line2.BorderColor
End Property

Public Property Let BorderColor2(ByVal New_BorderColor2 As OLE_COLOR)
    Line2.BorderColor() = New_BorderColor2
    PropertyChanged "BorderColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line3,Line3,-1,BorderColor
Public Property Get BorderColor3() As OLE_COLOR
Attribute BorderColor3.VB_Description = "Returns/sets the color of an object's border."
    BorderColor3 = Line3.BorderColor
End Property

Public Property Let BorderColor3(ByVal New_BorderColor3 As OLE_COLOR)
    Line3.BorderColor() = New_BorderColor3
    PropertyChanged "BorderColor3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line4,Line4,-1,BorderColor
Public Property Get BorderColor4() As OLE_COLOR
Attribute BorderColor4.VB_Description = "Returns/sets the color of an object's border."
    BorderColor4 = Line4.BorderColor
End Property

Public Property Let BorderColor4(ByVal New_BorderColor4 As OLE_COLOR)
    Line4.BorderColor() = New_BorderColor4
    PropertyChanged "BorderColor4"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line5,Line5,-1,BorderColor
Public Property Get BorderColor5() As OLE_COLOR
Attribute BorderColor5.VB_Description = "Returns/sets the color of an object's border."
    BorderColor5 = Line5.BorderColor
End Property

Public Property Let BorderColor5(ByVal New_BorderColor5 As OLE_COLOR)
    Line5.BorderColor() = New_BorderColor5
    PropertyChanged "BorderColor5"
End Property

