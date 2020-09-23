VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Reserved 
   BorderStyle     =   0  'None
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox bgMain 
      Height          =   6435
      Left            =   0
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   387
      TabIndex        =   13
      Top             =   0
      Width           =   5865
      Begin VB.TextBox txtcontactnoCode 
         BackColor       =   &H8000000B&
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "N/A"
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtcontactno 
         BackColor       =   &H8000000B&
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "N/A"
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ComboBox cbo_nationality 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_Reservation.frx":0000
         Left            =   1920
         List            =   "frm_Reservation.frx":001C
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Filipino"
         Top             =   3960
         Width           =   3615
      End
      Begin VB.TextBox txthomeadd 
         BackColor       =   &H8000000B&
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "N/A"
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox txtemail 
         BackColor       =   &H8000000B&
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "N/A"
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox txtcompany 
         BackColor       =   &H8000000B&
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "N/A"
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox txtmname 
         BackColor       =   &H8000000B&
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "N/A"
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtfname 
         BackColor       =   &H8000000B&
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "N/A"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtlname 
         BackColor       =   &H8000000B&
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "N/A"
         Top             =   1680
         Width           =   3615
      End
      Begin HMS.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   609
         BackColor       =   12735512
         Caption         =   "Manage Reservation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   9.75
         ForeColor       =   8421504
         GradTheme       =   2
      End
      Begin MSComctlLib.ImageList ilRecordIco 
         Left            =   360
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Reservation.frx":0070
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HMS.b8SContainer pbBGButton 
         Height          =   870
         Left            =   0
         TabIndex        =   25
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1535
         BorderColor     =   14215660
         Begin VB.Label lblroomno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   2280
            TabIndex        =   31
            Top             =   120
            Width           =   45
         End
         Begin VB.Label lblFloor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   2280
            TabIndex        =   30
            Top             =   360
            Width           =   45
         End
         Begin VB.Label lblRT 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   2280
            TabIndex        =   29
            Top             =   600
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room Tariff:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   28
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Floor No  :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   27
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room No :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   26
            Top             =   120
            Width           =   840
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   120
            Picture         =   "frm_Reservation.frx":060A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   720
         End
      End
      Begin HMS.b8SContainer b8SContainer1 
         Height          =   870
         Left            =   0
         TabIndex        =   32
         Top             =   5520
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1535
         BorderColor     =   14215660
         Begin HMS.b8Line b8Line1 
            Height          =   60
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   106
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   405
            Left            =   4440
            TabIndex        =   33
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   714
            Caption         =   "&Cancel"
            CapAlign        =   2
            BackStyle       =   2
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
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   405
            Left            =   3360
            TabIndex        =   12
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   714
            Caption         =   "&Save"
            CapAlign        =   2
            BackStyle       =   2
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
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H cmdAdd 
            Height          =   405
            Left            =   2280
            TabIndex        =   0
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   714
            Caption         =   "&Add"
            CapAlign        =   2
            BackStyle       =   2
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
      Begin MSComCtl2.DTPicker dtpChkOdate 
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   5040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   55705600
         CurrentDate     =   38139
      End
      Begin MSComCtl2.DTPicker dtReserve 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   4680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   55705600
         CurrentDate     =   39613
      End
      Begin VB.Label lblsep 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   36
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Check-In Date:"
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
         Left            =   120
         TabIndex        =   35
         Top             =   4740
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check-Out Date:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   5100
         Width           =   1605
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   4320
         Width           =   1485
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No/Mobile No:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   3240
         Width           =   1740
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name / Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Fill out Customer information!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   3540
      End
   End
End
Attribute VB_Name = "frm_Reserved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cBox As New CAutoCompleteComboBox
Dim rs As New ADODB.Recordset
Dim rsChkRes As New ADODB.Recordset


Private Sub cbo_nationality_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub cmdAdd_Click()
    UnlockDataEntry
End Sub

Private Sub cmdCancel_Click()
    mdiMain.Enabled = True
    Unload Me
End Sub

Private Sub cmdSave_Click()
    chkdata
End Sub


Private Sub dtpChkOdate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub dtReserve_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub Form_Activate()
    Me.dtReserve.Value = Now
    Me.dtpChkOdate.Value = Now
End Sub

Private Sub Form_Load()
'SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
'SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    reposition
    cBox.Init Me.cbo_nationality
End Sub

Private Sub reposition()
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdiMain.Enabled = True
End Sub

Private Sub pbBGButton_GotFocus()
    mdiMain.Enabled = False
End Sub

Private Sub txtcompany_GotFocus()
    HLTxt txtcompany
End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            Me.txtcontactnoCode.SetFocus
    End Select
End Sub

Private Sub txtcontactno_GotFocus()
    HLTxt txtcontactno
End Sub

Private Sub txtcontactno_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 46
            KeyAscii = 0
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcontactnoCode_Change()
 If Len(Me.txtcontactnoCode.Text) = 3 Then Me.txtcontactno.SetFocus
End Sub

Private Sub txtcontactnoCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 46
            KeyAscii = 0
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtemail_GotFocus()
    HLTxt txtemail
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtfname_GotFocus()
    HLTxt txtfname
End Sub

Private Sub txtfname_KeyPress(KeyAscii As Integer)
If Me.txtfname.Text = "N/A" Or txtfname.Text = "" Then
    Select Case KeyAscii
        Case vbKeySpace
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End If
End Sub

Private Sub txthomeadd_GotFocus()
    HLTxt txthomeadd
End Sub

Private Sub txthomeadd_KeyPress(KeyAscii As Integer)
If Me.txthomeadd.Text = "N/A" Or txthomeadd.Text = "" Then
    Select Case KeyAscii
        Case vbKeySpace
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End If
End Sub

Private Sub txtlname_GotFocus()
    HLTxt txtlname
End Sub

Private Sub txtlname_KeyPress(KeyAscii As Integer)
If Me.txtlname.Text = "N/A" Or txtlname.Text = "" Then
    Select Case KeyAscii
        Case vbKeySpace
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End If
End Sub

Private Sub txtmname_GotFocus()
    HLTxt txtmname
End Sub

Private Sub chkdata()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblCustomerInfo", cn, adOpenKeyset, adLockPessimistic
    
    If rsChkRes.State = adStateOpen Then rsChkRes.Close
    rsChkRes.Open "Select * from tblReserved", cn, adOpenKeyset, adLockPessimistic
    
    Do While rsChkRes.EOF = False
        If Format(Me.dtReserve.Value, "MMDDYY") >= Format(rsChkRes.Fields("rDate").Value, "MMDDYY") And Format(Me.dtReserve.Value, "MMDDYY") <= Format(rsChkRes.Fields("uDate").Value, "MMDDYY") And Me.lblroomno.Caption = rsChkRes.Fields("rNo").Value Then
            MsgBox "Selected Room already Reserved for that day." & vbCrLf & "Please Select another day or another room." & vbCrLf & vbCrLf & "Thank you", vbInformation
            Exit Sub
        ElseIf Format(Me.dtpChkOdate.Value, "MMDDYY") >= Format(rsChkRes.Fields("rDate").Value, "MMDDYY") And Format(Me.dtpChkOdate.Value, "MMDDYY") <= Format(rsChkRes.Fields("uDate").Value, "MMDDYY") And Me.lblroomno.Caption = rsChkRes.Fields("rNo").Value Then
            MsgBox "Selected Room already Reserved for that day." & vbCrLf & "Please Select another day or another room." & vbCrLf & vbCrLf & "Thank you", vbInformation
            Exit Sub
        ElseIf Format(Me.dtReserve.Value, "MMDDYY") <= Format(rsChkRes.Fields("rDate").Value, "MMDDYY") And Format(Me.dtpChkOdate.Value, "MMDDYY") >= Format(rsChkRes.Fields("uDate").Value, "MMDDYY") And Me.lblroomno.Caption = rsChkRes.Fields("rNo").Value Then
            MsgBox "Selected Room already Reserved for that day." & vbCrLf & "Please Select another day or another room." & vbCrLf & vbCrLf & "Thank you", vbInformation
        Exit Sub
        End If
        rsChkRes.MoveNext
    Loop
    
    If txtlname.Text = "N/A" Or txtlname.Text = "" Then
        MsgBox "Please enter Customer's Last Name. Thank You", vbInformation
        txtlname.SetFocus
        Exit Sub
    End If
    
    If txtfname.Text = "N/A" Or txtfname.Text = "" Then
        MsgBox "Please enter Customer's First Name. Thank You", vbInformation
        txtfname.SetFocus
        Exit Sub
    End If
    
    If txtmname.Text = "N/A" Or txtmname.Text = "" Then
        MsgBox "Please enter Customer's Middle Name. Thank You", vbInformation
        txtmname.SetFocus
        Exit Sub
    End If
    
    If txthomeadd.Text = "N/A" Or txthomeadd.Text = "" Then
        MsgBox "Please enter Customer's Home Address. Thank You", vbInformation
        txthomeadd.SetFocus
        Exit Sub
    End If
    With rs
        .AddNew
        .Fields("LastName").Value = Me.txtlname.Text
        .Fields("FirstName").Value = Me.txtfname.Text
        .Fields("MiddleName").Value = Me.txtmname.Text
        .Fields("Company").Value = Me.txtcompany.Text
        If Me.txtcontactnoCode.Text = "N/A" Then
            .Fields("ContactNo").Value = Me.txtcontactnoCode.Text
        Else
            .Fields("ContactNo").Value = Me.txtcontactnoCode.Text & "-" & Me.txtcontactno.Text
        End If
        .Fields("Email").Value = Me.txtemail.Text
        .Fields("Nationality").Value = Me.cbo_nationality.Text
        .Fields("HomeAdd").Value = Me.txthomeadd.Text
        .Fields("chkOdate").Value = Me.dtpChkOdate.Value
        .Fields("Room_no").Value = Me.lblroomno.Caption
        .Fields("Floor_no").Value = Me.lblFloor.Caption
        .Fields("Room_Tariff").Value = Me.lblRT.Caption
        .Fields("occu_date").Value = (FormatDateTime(Now, vbShortDate))
        .Fields("occu_time").Value = FormatDateTime(Now, vbShortTime)
        .Fields("Stat").Value = "R"
        .Fields("coDate").Value = "0"
        .Fields("coTime").Value = "0:00"
        .Update
    End With
    'frm_Confirm.Show
    
    
    get_trans_id
    updateRooms
    PassToPayment
    LockDataEntry
    MsgBox "Customer's Data successfully save." & vbCrLf & "Reservation Successful" & vbCrLf & vbCrLf & "Thank you", vbInformation
    frm_Welcome.loadRooms
    mdiMain.Enabled = True
    Unload Me
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub UnlockDataEntry()
    Me.txtlname.Locked = False
    Me.txtlname.BackColor = vbWhite
    Me.txtfname.Locked = False
    Me.txtfname.BackColor = vbWhite
    Me.txtmname.Locked = False
    Me.txtmname.BackColor = vbWhite
    Me.txtcompany.Locked = False
    Me.txtcompany.BackColor = vbWhite
    Me.txtcontactno.Locked = False
    Me.txtcontactno.BackColor = vbWhite
    Me.txtcontactnoCode.Locked = False
    Me.txtcontactnoCode.BackColor = vbWhite
    Me.txtemail.Locked = False
    Me.txtemail.BackColor = vbWhite
    Me.cbo_nationality.Locked = False
    Me.cbo_nationality.BackColor = vbWhite
    Me.txthomeadd.Locked = False
    Me.txthomeadd.BackColor = vbWhite
    Me.dtpChkOdate.Enabled = True
    Me.lblsep.BackColor = vbWhite
    Me.cmdAdd.Enabled = False
    Me.cmdSave.Enabled = True
    Me.dtReserve.Enabled = True
End Sub

Private Sub LockDataEntry()
    Me.txtlname.Locked = True
    Me.txtlname.BackColor = &H8000000B
    Me.txtlname.Text = "N/A"
    Me.txtfname.Locked = True
    Me.txtfname.BackColor = &H8000000B
    Me.txtfname.Text = "N/A"
    Me.txtmname.Locked = True
    Me.txtmname.BackColor = &H8000000B
    Me.txtmname.Text = "N/A"
    Me.txtcompany.Locked = True
    Me.txtcompany.BackColor = &H8000000B
    Me.txtcompany.Text = "N/A"
    Me.txtcontactnoCode.Locked = True
    Me.txtcontactnoCode.BackColor = &H8000000B
    Me.txtcontactno.Locked = True
    Me.txtcontactno.BackColor = &H8000000B
    Me.txtcontactno.Text = "N/A"
    Me.txtemail.Locked = True
    Me.txtemail.BackColor = &H8000000B
    Me.txtemail.Text = "N/A"
    Me.cbo_nationality.Locked = True
    Me.cbo_nationality.BackColor = &H8000000B
    Me.txthomeadd.Locked = True
    Me.txthomeadd.BackColor = &H8000000B
    Me.txthomeadd.Text = "N/A"
    Me.dtpChkOdate.Enabled = False
    Me.dtReserve.Enabled = False
    Me.cmdAdd.Enabled = True
    Me.cmdSave.Enabled = False
End Sub

Public Sub updateRooms()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblReserved", cn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs.Fields("rNo").Value = Me.lblroomno.Caption
    rs.Fields("TransID").Value = transID
    rs.Fields("rDate").Value = FormatDateTime(Me.dtReserve.Value, vbShortDate)
    rs.Fields("rTime").Value = FormatDateTime(Now, vbShortTime)
    rs.Fields("uDate").Value = FormatDateTime(Me.dtpChkOdate.Value, vbShortDate)
    rs.Update
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub get_trans_id()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblCustomerInfo", cn, adOpenKeyset, adLockPessimistic
    rs.MoveLast
    transID = rs.Fields("TransactionID").Value
    'frm_Welcome.lvr.SelectedItem.SubItems(5) = transID
    'rs.Update
    Exit Sub
err:
    MsgBox err.Description, vbCritical

End Sub
Private Sub PassToPayment()
On Error GoTo err:
If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblPayment", cn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs.Fields("TransactionID").Value = transID
    rs.Fields("RoomNo").Value = frm_Welcome.lvr.SelectedItem.SubItems(4)
    'rs.Fields("Stat").Value = "true"
    rs.Update
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub txtmname_KeyPress(KeyAscii As Integer)
If Me.txtmname.Text = "N/A" Or txtmname.Text = "" Then
    Select Case KeyAscii
        Case vbKeySpace
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End If
End Sub
