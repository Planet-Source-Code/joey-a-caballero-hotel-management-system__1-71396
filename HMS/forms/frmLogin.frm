VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User's Login"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2970
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2970
      Width           =   1395
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1740
      Width           =   3405
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "="
      TabIndex        =   3
      Top             =   2280
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Picture         =   "frmLogin.frx":058A
      Top             =   0
      Width           =   4830
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   1785
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim passattemp As Integer
Dim currentemmsuser As String
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdLogin_Click()
    On Error GoTo err

'-Verify the fields if empty
If txtUsername.Text = "" Then txtUsername.SetFocus: Exit Sub
If txtPassword.Text = "" Then txtPassword.SetFocus: Exit Sub

'-Check if the User Name is valid

        With rs

        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM tblUser WHERE username = """ & txtUsername.Text & _
        """ and Password = """ & txtPassword.Text & """", cn, adOpenKeyset, adLockOptimistic
        
        If .EOF = False Then
            currentemmsuser = .Fields("Fullname")
            'sndPlaySound App.Path & "\mssound.wav", 0 + 17
            'frm_Main.Show
            'frm_Main.Visible = True
            lpass = txtPassword.Text
            With mdiMain
                .lblusername.Caption = currentemmsuser
                
                If rs.Fields("Priv").Value = "Cashier" Then
                    .mnuMaintenance.Visible = False
                    .mnuUser.Visible = False
                    mdiMain.loadQuickLauchC
                Else
                    mdiMain.loadQuickLauch
                    .mnuMaintenance.Visible = True
                    .mnuUser.Visible = True
                    '.mnuUsersLog.Visible = False
                End If
            End With
            sLog
            Unload Me
            mdiMain.Show
        Else
            passattemp = passattemp + 1
            If passattemp = 3 Then
           'sndPlaySound App.Path & "\Mad.wav", 0 + 17
                MsgBox "You are not an authorized user", vbInformation + vbCritical, "Log In Error"
                End
            End If
            Me.txtUsername.SetFocus
            MsgBox "Invalid Username or Password" & vbCrLf & "ACCESS DENIED!!!" & vbCrLf & vbCrLf & "Please try again." & vbCrLf & vbCrLf, vbCritical, "Log In Error"
            Me.txtPassword.SetFocus
            Exit Sub
        End If
        .Close

            'Dim strSql As String
            'strSql = "Select * From tblUser Where username='" & txtUserName.Text & "'"
            '.Open strSql, cn, adOpenStatic, adLockOptimistic
            'MsgBox .RecordCount
            'If .RecordCount >= 1 Then
            '    If .Fields("Password") = txtPassword.Text Then
            '        lpass = txtPassword.Text
            '        currentemmsuser = .Fields("Fullname")
            '        'CurrentPosition = .Fields("USER_TYPE")
             '           'sndPlaySound App.Path & "\mssound.wav", 0 + 17
             '           With mdiMain
             ''               .lblusername.Caption = currentemmsuser
             '
             '           End With
             '           Unload Me
              '          mdiMain.Show
       '
       '         Else
       '             passattemp = passattemp + 1
       '             If passattemp = 3 Then
       '                 'sndPlaySound App.Path & "\Mad.wav", 0 + 17
       '                 MsgBox "You are not an authorized user", vbInformation + vbCritical, "Log In Error"
       '                 End
       '             Else
       '                 MsgBox "Password incorrect. Please check the CAPS LOCK" & vbCrLf & " Attempt left " & 3 - passattemp & "", vbExclamation, "Log In Error"
        ''                txtPassword.Text = ""
        '                txtPassword.SetFocus
        '            End If
        '        End If
        '    Else
        '        MsgBox "This user does not exist", vbCritical, "Log In Error"
        '        txtUserName.Text = ""
        '        txtUserName.SetFocus
        '    End If
        '    .Close
        End With
        
Set rs = Nothing

Exit Sub

err:
MsgBox err.Description, vbCritical
End Sub

Private Sub sLog()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tbluserLog", cn, adOpenKeyset, adLockPessimistic
        rs.AddNew
        rs.Fields("Fullname").Value = currentemmsuser
        rs.Fields("username").Value = Me.txtUsername.Text
        rs.Fields("oDate").Value = FormatDateTime(Now, vbShortDate)
        rs.Fields("oTime").Value = FormatDateTime(Now, vbLongTime)
        rs.Update
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Load()
    qL = True
End Sub

Private Sub txtPassword_GotFocus()
    HLTxt txtPassword
End Sub

Private Sub txtUsername_GotFocus()
    HLTxt txtUsername
End Sub

