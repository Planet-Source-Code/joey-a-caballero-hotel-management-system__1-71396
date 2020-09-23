VERSION 5.00
Begin VB.UserControl b8Container 
   BackColor       =   &H00F6F8F8&
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   ControlContainer=   -1  'True
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   Begin VB.Shape ibr 
      BackColor       =   &H00BFCED0&
      BorderColor     =   &H00808080&
      Height          =   4215
      Left            =   180
      Top             =   240
      Width           =   3015
   End
   Begin VB.Shape bri2 
      BackColor       =   &H00BFCED0&
      BorderColor     =   &H00FFFFFF&
      Height          =   4215
      Left            =   60
      Top             =   810
      Width           =   3015
   End
   Begin VB.Shape br2 
      BackColor       =   &H00BFCED0&
      BorderColor     =   &H00D0E0E3&
      Height          =   4215
      Left            =   0
      Top             =   270
      Width           =   3675
   End
   Begin VB.Shape br1 
      BackColor       =   &H00BFCED0&
      BorderColor     =   &H00BFCED0&
      Height          =   4215
      Left            =   240
      Top             =   60
      Width           =   3675
   End
End
Attribute VB_Name = "b8Container"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'code by: Vincent Jamero

Option Explicit



Private Sub UserControl_Resize()
On Error Resume Next
    br2.Move 0, 0, GetWidth - 0, GetHeight - 0
    br1.Move 1, 1, GetWidth - 2, GetHeight - 2
    ibr.Move 2, 2, GetWidth - 4, GetHeight - 4
    bri2.Move 3, 3, GetWidth - 6, GetHeight - 6

End Sub




Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function
    
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ibr,ibr,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = ibr.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    ibr.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As ScaleModeConstants)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ibr.BorderColor = PropBag.ReadProperty("BorderColor", &H808080)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 3)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HD8E9EC)
    bri2.BorderColor = PropBag.ReadProperty("InsideBorderColor", 16777215)
    br1.BorderColor = PropBag.ReadProperty("ShadowColor1", 12570320)
    br2.BorderColor = PropBag.ReadProperty("ShadowColor2", 13689059)
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderColor", ibr.BorderColor, 8421504)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 3)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HD8E9EC)
    Call PropBag.WriteProperty("InsideBorderColor", bri2.BorderColor, 16777215)
    Call PropBag.WriteProperty("ShadowColor1", br1.BorderColor, 12570320)
    Call PropBag.WriteProperty("ShadowColor2", br2.BorderColor, 13689059)
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
'MappingInfo=bri2,bri2,-1,BorderColor
Public Property Get InsideBorderColor() As OLE_COLOR
Attribute InsideBorderColor.VB_Description = "Returns/sets the color of an object's border."
    InsideBorderColor = bri2.BorderColor
End Property

Public Property Let InsideBorderColor(ByVal New_InsideBorderColor As OLE_COLOR)
    bri2.BorderColor() = New_InsideBorderColor
    PropertyChanged "InsideBorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=br1,br1,-1,BorderColor
Public Property Get ShadowColor1() As OLE_COLOR
Attribute ShadowColor1.VB_Description = "Returns/sets the color of an object's border."
    ShadowColor1 = br1.BorderColor
End Property

Public Property Let ShadowColor1(ByVal New_ShadowColor1 As OLE_COLOR)
    br1.BorderColor() = New_ShadowColor1
    PropertyChanged "ShadowColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=br2,br2,-1,BorderColor
Public Property Get ShadowColor2() As OLE_COLOR
Attribute ShadowColor2.VB_Description = "Returns/sets the color of an object's border."
    ShadowColor2 = br2.BorderColor
End Property

Public Property Let ShadowColor2(ByVal New_ShadowColor2 As OLE_COLOR)
    br2.BorderColor() = New_ShadowColor2
    PropertyChanged "ShadowColor2"
End Property

