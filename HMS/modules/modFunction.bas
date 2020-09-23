Attribute VB_Name = "modFunction"
Option Explicit


'functions
Public Enum FindOptions
    PartOfWord = 0
    MatchCase = 1
    WholeWordOnly = 3
End Enum


'API for opening a browser
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long


Public Sub FormDrag(frmName As Form) 'procedure to drag a no-titlebar form
    ReleaseCapture
    Call SendMessage(frmName.hwnd, &HA1, 2, 0&)
End Sub



Public Function MakeGradient(ByRef Frm As Object, Scheme As Integer)
    Dim cr(255) As Integer
    Dim cG(255) As Integer
    Dim cB(255) As Integer
    Dim d As Double
    Dim i As Integer
    
    
    Select Case Scheme
        Case 1
            For i = 0 To 255
                cr(i) = 255 - (i * 0.2)
                cG(i) = 255 - (i * 0.2)
                cB(i) = 255 - (i * 0.2)
            Next
    End Select
    

    Frm.ScaleMode = vbPixels
    d = Frm.ScaleHeight / 255
    Frm.DrawWidth = d + 1
    For i = 0 To 255
        Frm.ForeColor = RGB(cr(i), cG(i), cB(i))
        Frm.Line (0, i * d)-(Frm.ScaleWidth, i * d)
    Next
    'Frm.AutoRedraw = True
End Function

Public Function CheckTextBox(ByRef txt As Object, Optional sMSG As String = "TextBox", Optional ShowMSG As Boolean = True, Optional MinimumChar As Integer = 1) As Boolean
On Error Resume Next
    If Len(Trim(txt.Text)) < MinimumChar Then
        
        If ShowMSG Then
            MsgBox sMSG, vbExclamation
        End If
        
        txt.Text = ""
        txt.SetFocus
        
        CheckTextBox = False
    Else
        CheckTextBox = True
    End If
End Function

Public Function HLTxt(ByRef txt As Object)
On Error Resume Next
    txt.SelStart = 0
    txt.SelLength = Len(txt)
    txt.SetFocus
End Function


Public Function CatchError(sModuleName As String, sRoutineName As String, sDetail As String)
    MsgBox sModuleName & " - " & sRoutineName & " - " & sDetail
    
End Function


Public Function CenterForm(ByRef Frm As Form)
    Frm.Move (Screen.Width - Frm.Width) / 2, (Screen.Height - Frm.Height) / 2
End Function

Public Function cWords(ByVal strTheString As String) As String
    'Description: Capitalize the first letter of each word in a string
    Dim cr As String
    Dim t As String
    Dim i
    cr = Chr$(13) + Chr$(10)
    t = strTheString  'the string
    If t <> "" Then
        Mid$(t, 1, 1) = UCase$(Mid$(t, 1, 1))
        For i = 1 To Len(t$) - 1
            If Mid$(t, i, 2) = cr Then Mid$(t$, i + 2, 1) = UCase$(Mid$(t, i + 2, 1))
            If Mid$(t, i, 1) = " " Then Mid$(t$, i + 1, 1) = UCase$(Mid$(t, i + 1, 1))
        Next
        cWords = t
    End If
End Function
