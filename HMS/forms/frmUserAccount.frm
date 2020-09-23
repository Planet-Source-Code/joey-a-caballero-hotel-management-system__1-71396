VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUserAccount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   620
   StartUpPosition =   2  'CenterScreen
   Begin HMS.b8SContainer b8SContainer1 
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   5520
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   926
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   405
         Left            =   7560
         TabIndex        =   2
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         Caption         =   "&Close"
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
   End
   Begin VB.Timer timerUSU 
      Interval        =   1
      Left            =   4050
      Top             =   2790
   End
   Begin MSComctlLib.ImageList ilRecordIcos 
      Left            =   3870
      Top             =   3510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":058A
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":0B24
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilRecordIco 
      Left            =   3810
      Top             =   2820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":10BE
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAccount.frx":1F98
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listRecord 
      Height          =   4935
      Left            =   2430
      TabIndex        =   0
      Top             =   570
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8705
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ilRecordIcos"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmUserAccount.frx":2E72
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User Type"
         Object.Width           =   2699
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdAdd 
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Add"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdEdit 
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   3990
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Edit"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   4380
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Delete"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin lvButton.lvButtons_H cmdReload 
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   4770
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      Caption         =   "&Reload"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin HMS.b8Line b8Line1 
      Height          =   60
      Left            =   690
      TabIndex        =   7
      Top             =   510
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   106
   End
   Begin lvButton.lvButtons_H cmdViewLog 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Caption         =   "&View Users Log Record"
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   8421504
      cFHover         =   0
      cBhover         =   16777215
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   4
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin VB.Label lblFullName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00808080&
      Height          =   555
      Left            =   150
      TabIndex        =   13
      Top             =   1710
      Width           =   2235
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   1500
      Width           =   2235
   End
   Begin VB.Label lblUserName 
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
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1290
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   90
      TabIndex        =   10
      Top             =   960
      Width           =   2085
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "frmUserAccount.frx":374C
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Accounts"
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
      Left            =   720
      TabIndex        =   8
      Top             =   180
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -30
      Picture         =   "frmUserAccount.frx":4016
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   0
      Picture         =   "frmUserAccount.frx":4EE0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9585
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F6F8F8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   5085
      Left            =   0
      Top             =   510
      Width           =   2430
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowForm()
    
    'check acess/admin
    If CurrentUser.UserType <> sAdministratortitle Then
        MsgBox "Unable to show Manage Users window." & vbNewLine & _
                "You are not permitted to aceess it. Please contact your Administrator.", vbExclamation
                
        Unload Me
        Exit Sub
    End If
    
    'set access/controls
    CheckAccessControls
    Form_FillList
    
    Me.Show vbModal
End Sub

                        
Private Function CheckAccessControls()

    cmdAdd.Enabled = UserAllowedTo(CurrentUser.UserName, sCanAddUser)
    cmdEdit.Enabled = UserAllowedTo(CurrentUser.UserName, sCanEditUser)
    cmdDelete.Enabled = UserAllowedTo(CurrentUser.UserName, sCanDeleteUser)
    
End Function


Private Sub cmdAdd_Click()
    frmAddUser.ShowForm
    Form_FillList

End Sub


Private Sub Form_FillList()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    
    sSQL = "SELECT tblUser.UserName, tblUser.UserName, tblUser.FullName, tblUser.UserType" & _
            " FROM tblUser;"
    
    'clear list items
    listRecord.ListItems.Clear
    
    If ConnectRS(HSESDB, vRS, sSQL) = True Then
        If AnyRecordExisted(vRS) = False Then
            listRecord.Enabled = False
            GoTo ReleaseAndExit
        End If
    Else
        GoTo ReleaseAndExit
        CatchError "frmUserAccount", "Form_FillList", "ConnectRS"
    End If
    
    vRS.MoveFirst
    While vRS.EOF = False
    
        If LCase(ReadField(vRS.Fields("UserType"))) = "administrator" Then
            listRecord.ListItems.Add , keyUser & ReadField(vRS.Fields("UserName")), ReadField(vRS.Fields("UserName")), "admin"
        Else
            listRecord.ListItems.Add , keyUser & ReadField(vRS.Fields("UserName")), ReadField(vRS.Fields("UserName")), "user"
        End If
    
        listRecord.ListItems(listRecord.ListItems.Count).SubItems(1) = ReadField(vRS.Fields("FullName"))
        listRecord.ListItems(listRecord.ListItems.Count).SubItems(2) = ReadField(vRS.Fields("UserType"))

        
        vRS.MoveNext
    Wend
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


Private Sub cmdCLose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim lvKey As String
    
    lvKey = GetLVKey(listRecord.SelectedItem)

    If lvKey = CurrentUser.UserName Then
        MsgBox "Unable to continue deleting User entry." & vbNewLine & _
        "You cannot delete your own record.", vbExclamation
        Exit Sub
    End If
    
    If "administrator" = LCase(lvKey) Then
        MsgBox "Unable to continue deleting User entry." & vbNewLine & _
        "You cannot delete the Head Administrator entry.", vbExclamation
        Exit Sub
    End If
    
    If DeleteUser(lvKey) = Success Then
        MsgBox "Entry deleted.", vbInformation
        Form_FillList
    Else
        MsgBox "Unable to delete user entry.", vbExclamation
    End If
    
End Sub

Private Sub cmdEdit_Click()
    Dim lvKey As String
    
    lvKey = GetLVKey(listRecord.SelectedItem)
    
    If lvKey = CurrentUser.UserName Then
        MsgBox "Unable to continue editing User entry." & vbNewLine & _
        "You cannot modify your own record.", vbExclamation
        Exit Sub
    End If
    
    If "administrator" = LCase(lvKey) Then
        MsgBox "Unable to continue editing User entry." & vbNewLine & _
        "You cannot modifiy the Head Administrator entry.", vbExclamation
        Exit Sub
    End If
    
    If Len(lvKey) > 0 Then
        frmEditUser.ShowForm (lvKey)
        Form_FillList

    Else
        MsgBox "Please select User", vbExclamation
    End If
End Sub

Private Sub cmdReload_Click()
    Form_FillList
End Sub

Private Sub cmdViewLog_Click()
    frmUserLog.ShowForm
End Sub

Private Sub timerUSU_Timer()
    Static sName As String
    
    If sName = listRecord.SelectedItem.Text Then
        Exit Sub
    End If
    
    lblUserName.Caption = listRecord.SelectedItem.Text
    lblType.Caption = listRecord.SelectedItem.SubItems(2)
    lblFullName.Caption = listRecord.SelectedItem.SubItems(1)
    sName = listRecord.SelectedItem.Text
    
End Sub
