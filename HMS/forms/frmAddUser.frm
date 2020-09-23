VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAddUser 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtConfirm 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   3540
   End
   Begin VB.ComboBox cboPriv 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmAddUser.frx":058A
      Left            =   1200
      List            =   "frmAddUser.frx":0594
      TabIndex        =   5
      Text            =   "Select Privelege"
      Top             =   2640
      Width           =   3540
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1170
      Width           =   3540
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   3540
   End
   Begin VB.TextBox txtFullName 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   1
      Top             =   750
      Width           =   3540
   End
   Begin HMS.b8Line b8Line1 
      Height          =   60
      Left            =   -30
      TabIndex        =   10
      Top             =   3120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   106
   End
   Begin HMS.b8SContainer b8SContainer1 
      Height          =   570
      Left            =   -60
      TabIndex        =   11
      Top             =   5280
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1005
      BorderColor     =   14737632
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   4080
         TabIndex        =   9
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Caption         =   "&Cancel"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14215660
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   405
         Left            =   3135
         TabIndex        =   8
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Caption         =   "&Delete"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14215660
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   405
         Left            =   2175
         TabIndex        =   6
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Caption         =   "&Save"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14215660
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   405
         Left            =   1215
         TabIndex        =   7
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Caption         =   "&Edit"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14215660
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdAdd 
         Default         =   -1  'True
         Height          =   405
         Left            =   255
         TabIndex        =   0
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Caption         =   "&Add"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   14215660
         LockHover       =   1
         cGradient       =   14215660
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
   End
   Begin MSComctlLib.ListView listAccess 
      Height          =   1905
      Left            =   60
      TabIndex        =   12
      Top             =   3330
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3360
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Full Name"
         Object.Width           =   6085
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Privilege"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   0
      EndProperty
   End
   Begin HMS.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   16
      Top             =   510
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   106
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   19
      Top             =   2220
      Width           =   810
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Privilege:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   18
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   180
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   15
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   14
      Top             =   1740
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   13
      Top             =   750
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmAddUser.frx":05B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5145
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isEdit As Boolean
Dim rsUsers As ADODB.Recordset
Dim cboP As New CAutoCompleteComboBox

Private Sub cmdAdd_Click()
    isEdit = False
    Me.txtFullName.Enabled = True
    Me.txtPassword.Enabled = True
    Me.txtUsername.Enabled = True
    Me.txtConfirm.Enabled = True
    Me.cboPriv.Enabled = True
    
    Me.txtFullName.Text = ""
    Me.txtPassword.Text = ""
    Me.txtUsername.Text = ""
    Me.cboPriv.Text = ""
    Me.txtConfirm.Text = ""
    Me.txtFullName.SetFocus
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo eh
    
    Dim intYN
    
    intYN = MsgBox("You are about to delete a record." & vbCrLf & _
        "If you click Yes, you won't be able to undo this delete operation." & _
        vbCrLf & vbCrLf & _
        "Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm Delete")
        
    If intYN = vbNo Then Exit Sub
    
    cn.Execute "DELETE FROM tblUser WHERE ID = " & Me.listAccess.SelectedItem.SubItems(4)
    'Me.lvwUser.SetFocus
    'Me.lvwUser.ListItems.Item(1).Selected = True
    'lvwUser_ItemClick
    Call GetUsers
    txtFullName.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtConfirm.Text = ""
    cboPriv.Text = ""
    MsgBox "Record deleted.", vbInformation

    Exit Sub
    
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub cmdEdit_Click()
    Me.txtFullName.Enabled = True
    Me.txtPassword.Enabled = True
    Me.txtUsername.Enabled = True
    Me.txtConfirm.Enabled = True
    Me.cboPriv.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    txtFullName.SetFocus
    isEdit = True
End Sub

Private Sub cmdSave_Click()
On Error GoTo eh
    
    If txtUsername.Text = vbNullString Then
        MsgBox "Please enter user name.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtFullName.Text = vbNullString Then
        MsgBox "Please enter Fullname.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtPassword.Text = vbNullString Then
        MsgBox "Please enter Password.", vbExclamation
        txtPass.SetFocus
        Exit Sub
    End If
    
    If txtConfirm.Text = vbNullString Then
        MsgBox "Please Confirm Password.", vbExclamation
        txtConfirm.SetFocus
        Exit Sub
    End If
    
    If cboPriv.Text = vbNullString Then
        MsgBox "Please Select Priviledge.", vbExclamation
        cboPriv.SetFocus
        Exit Sub
    End If
    
    If txtPassword.Text <> txtConfirm.Text Then
        MsgBox "Password does not match!", vbExclamation
        txtConfirm.SetFocus
        Exit Sub
    End If
        
    Dim rsSave As ADODB.Recordset
    
    Set rsSave = New ADODB.Recordset
            
    With rsSave
        If isEdit = False Then         'add record to database
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM tblUser WHERE ID = 0;", cn, adOpenKeyset, adLockOptimistic
            .AddNew
            .Fields("Fullname") = txtFullName.Text
            .Fields("username") = txtUsername.Text
            .Fields("password") = txtPassword.Text
            .Fields("Priv") = cboPriv.Text
            .Update
            MsgBox "New record saved.", vbInformation
        ElseIf isEdit = True Then    'update current record
            .Open "SELECT * FROM tblUser WHERE ID = " & Me.listAccess.SelectedItem.SubItems(4), cn, adOpenKeyset, adLockOptimistic
            .Fields("Fullname") = txtFullName.Text
            .Fields("username") = txtUsername.Text
            .Fields("password") = txtPassword.Text
            .Fields("Priv") = cboPriv.Text
            .Update
            MsgBox "Record updated.", vbInformation
        End If
        GetUsers
    End With
    cmdSave.Enabled = False
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
    txtFullName.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtConfirm.Text = ""
    cboPriv.Text = ""
    Exit Sub
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Load()
    GetUsers
    cboP.Init Me.cboPriv
End Sub

Private Sub listAccess_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo eh
    Set rsUsers = New ADODB.Recordset

    With rsUsers
        If .State = adStateOpen Then .Close
        '.CursorLocation = adUseClient
        .Open "SELECT * FROM tblUser where ID = " & Me.listAccess.SelectedItem.SubItems(4), cn, adOpenKeyset, adLockOptimistic
            Me.txtFullName.Text = .Fields("Fullname")
            Me.txtUsername.Text = .Fields("username")
            Me.txtPassword.Text = .Fields("password")
            Me.txtConfirm.Text = .Fields("password")
            Me.cboPriv.Text = .Fields("Priv")
    End With
    Exit Sub
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtFullName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub GetUsers()
    On Error GoTo eh
        
    Set rsUsers = New ADODB.Recordset

    With rsUsers
        If .State = adStateOpen Then .Close
        '.CursorLocation = adUseClient
        .Open "SELECT * FROM tblUser ORDER BY Username;", cn, adOpenKeyset, adLockOptimistic
        If rsUsers.EOF = True Then
            Me.listAccess.ListItems.Clear
            Exit Sub
        End If
        
        Me.listAccess.ListItems.Clear
        Do While .EOF = False
            Me.listAccess.ListItems.Add , , .Fields("Fullname")
            Me.listAccess.ListItems(Me.listAccess.ListItems.Count).SubItems(1) = .Fields("Username")
            Me.listAccess.ListItems(Me.listAccess.ListItems.Count).SubItems(2) = "*********"
            Me.listAccess.ListItems(Me.listAccess.ListItems.Count).SubItems(3) = .Fields("Priv")
            Me.listAccess.ListItems(Me.listAccess.ListItems.Count).SubItems(4) = .Fields("ID")
            .MoveNext
        Loop
    End With


    Exit Sub
    
eh:
    MsgBox err.Description, vbCritical
End Sub
