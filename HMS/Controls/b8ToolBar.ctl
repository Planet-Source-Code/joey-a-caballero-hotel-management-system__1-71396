VERSION 5.00
Begin VB.UserControl b8ToolBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   ControlContainer=   -1  'True
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   Begin VB.Line Line2 
      BorderColor     =   &H00CCE2E6&
      X1              =   -4
      X2              =   438
      Y1              =   41
      Y2              =   41
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00BBCACD&
      X1              =   7
      X2              =   449
      Y1              =   65
      Y2              =   65
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C3DADE&
      X1              =   -1
      X2              =   441
      Y1              =   55
      Y2              =   55
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D8E9EC&
      X1              =   13
      X2              =   455
      Y1              =   17
      Y2              =   17
   End
End
Attribute VB_Name = "b8ToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Event Declarations:
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."





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





'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub



'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub



Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelY
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelX
End Function



Private Sub UserControl_Resize()
    RaiseEvent Resize
    On Error Resume Next
    
    Line1.X1 = 0
    Line1.X2 = GetWidth
    Line1.Y1 = 0
    Line1.Y2 = 0
    
    Line2.X1 = 0
    Line2.X2 = GetWidth
    Line2.Y1 = GetHeight - 3
    Line2.Y2 = GetHeight - 3
    
    Line3.X1 = 0
    Line3.X2 = GetWidth
    Line3.Y1 = GetHeight - 2
    Line3.Y2 = GetHeight - 2
    
    Line4.X1 = 0
    Line4.X2 = GetWidth
    Line4.Y1 = GetHeight - 1
    Line4.Y2 = GetHeight - 1
End Sub

