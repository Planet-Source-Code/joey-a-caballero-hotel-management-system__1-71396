VERSION 5.00
Begin VB.UserControl b8SContainer 
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Shape shapeBorder 
      BorderColor     =   &H0030A0B8&
      Height          =   1410
      Left            =   1125
      Top             =   1245
      Width           =   3210
   End
   Begin VB.Image imgBGBottom 
      Height          =   300
      Left            =   15
      Picture         =   "bgSContainer.ctx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2085
   End
   Begin VB.Image imgBGTop 
      Height          =   90
      Left            =   0
      Picture         =   "bgSContainer.ctx":015A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2280
   End
End
Attribute VB_Name = "b8SContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Event Declarations:
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."


Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelY
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelX
End Function

Private Sub UserControl_Resize()
    RaiseEvent Resize
    On Error Resume Next
    
    imgBGTop.Move 0, 1, GetWidth
    imgBGBottom.Move 0, GetHeight - imgBGBottom.Height - 1, GetWidth

    shapeBorder.Move 0, 0, GetWidth, GetHeight
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shapeBorder,shapeBorder,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = shapeBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    shapeBorder.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    shapeBorder.BorderColor = PropBag.ReadProperty("BorderColor", 3186872)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderColor", shapeBorder.BorderColor, 3186872)
End Sub

