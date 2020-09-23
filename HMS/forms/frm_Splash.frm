VERSION 5.00
Begin VB.Form frm_Splash 
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   4590
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   5040
   End
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      MousePointer    =   11  'Hourglass
      Picture         =   "frm_Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer

Private Sub Timer1_Timer()
    'MsgBox Format(TimeSerial(h, m, s), "HH:MM:SS")
    x = x + 1
    If x = 15 Then
        'Unload Me
        Call dbconnect
    End If
End Sub
