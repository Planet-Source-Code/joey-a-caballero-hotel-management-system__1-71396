Attribute VB_Name = "modSystemColors"
Option Explicit


'Copyright by Philip V. Naparan | http://www.philipnaparan.cjb.net

'---------------------------------------------------
'PUT THIS IN THE MODULE
'---------------------------------------------------
Global original_menu_color              As Long
Global original_buttonface_color        As Long
Global original_buttonshadow_color      As Long
Global original_buttonhighlight_color      As Long

Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_2NDACTIVECAPTION = 27
Public Const COLOR_2NDINACTIVECAPTION = 28
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12


Public Type ColorSystem
    SelectColor(0 To 20) As Long
End Type

Public New_System_Color As ColorSystem


Public Sub select_color_type()
            'This is an XP Style color system
            New_System_Color.SelectColor(4) = RGB(239, 238, 224)  'Menu
            New_System_Color.SelectColor(15) = &HD8E9EC   'RGB(240, 240, 224) 'Button
            New_System_Color.SelectColor(16) = RGB(216, 210, 189) 'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight
            New_System_Color.SelectColor(COLOR_MENUTEXT) = RGB(0, 0, 0)
            New_System_Color.SelectColor(COLOR_BACKGROUND) = &HD8E9EC
            Call change_system_color
End Sub

Public Sub change_system_color()

Call SetSysColors(1, 4, New_System_Color.SelectColor(4))   'Menu
Call SetSysColors(1, 15, New_System_Color.SelectColor(15)) 'Button
Call SetSysColors(1, 16, New_System_Color.SelectColor(16)) 'Button Shadow
Call SetSysColors(1, 20, New_System_Color.SelectColor(20)) 'Button Highlight
Call SetSysColors(1, COLOR_MENUTEXT, New_System_Color.SelectColor(COLOR_MENUTEXT))
End Sub

'---------------------------------------------------
'END CODE FOR MODULE
'---------------------------------------------------


