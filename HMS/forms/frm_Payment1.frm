VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Payment1 
   BorderStyle     =   0  'None
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvlList 
      Height          =   735
      Left            =   600
      TabIndex        =   24
      Top             =   6480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   384
      ImageHeight     =   140
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Payment1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Payment1.frx":5284
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bgMain 
      Height          =   4035
      Left            =   0
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   347
      TabIndex        =   8
      Top             =   0
      Width           =   5265
      Begin VB.CheckBox chkAddPer 
         Caption         =   "Please Check this checkbox if Customer is Zero-Rated."
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
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   4815
      End
      Begin HMS.b8Line b8Line2 
         Height          =   60
         Left            =   0
         TabIndex        =   21
         Top             =   3120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   106
      End
      Begin VB.ComboBox cboType 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frm_Payment1.frx":A198
         Left            =   3480
         List            =   "frm_Payment1.frx":A1A5
         TabIndex        =   6
         Text            =   "Cash"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtRoomRate 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtAmountDue 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtPayment 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1320
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   1800
         Width           =   1215
      End
      Begin HMS.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   609
         BackColor       =   12735512
         Caption         =   "Manage Payment"
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
               Picture         =   "frm_Payment1.frx":A1BA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HMS.b8SContainer pbBGButton 
         Height          =   870
         Left            =   -120
         TabIndex        =   9
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1535
         BorderColor     =   14215660
         Begin VB.Label lblPaid 
            Caption         =   "0.00"
            Height          =   255
            Left            =   3840
            TabIndex        =   23
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblcBalance 
            Caption         =   "0.00"
            Height          =   255
            Left            =   3840
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   120
            Picture         =   "frm_Payment1.frx":A754
            Stretch         =   -1  'True
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Check-In:"
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
            TabIndex        =   13
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Check-In:"
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
            TabIndex        =   12
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label lblTime 
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
            Left            =   2520
            TabIndex        =   11
            Top             =   480
            Width           =   45
         End
         Begin VB.Label lbldate 
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
            Left            =   2520
            TabIndex        =   10
            Top             =   240
            Width           =   45
         End
      End
      Begin HMS.b8SContainer b8SContainer1 
         Height          =   870
         Left            =   0
         TabIndex        =   14
         Top             =   3120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1535
         BorderColor     =   14215660
         Begin HMS.b8Line b8Line1 
            Height          =   60
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   106
         End
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   405
            Left            =   3360
            TabIndex        =   7
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            Caption         =   "Cl&ose"
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
            Enabled         =   0   'False
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   405
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            Caption         =   "&Save Payment"
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
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Room Rate:"
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
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Due:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2640
         TabIndex        =   19
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment:"
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
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
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
         TabIndex        =   17
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
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
         Left            =   2640
         TabIndex        =   16
         Top             =   2160
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_Payment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rsClear As New ADODB.Recordset
Dim roomrate As Currency
Dim cbo_days As New CAutoCompleteComboBox

Private Sub b8Title_GotFocus()
 mdiMain.Enabled = False
End Sub

Private Sub bgMain_GotFocus()
 mdiMain.Enabled = False
End Sub

Private Sub cmdClose_Click()
    'MsgBox ischkout
    If isPayment = False Then
    'MsgBox isPayment
        SaveBills
    End If
    
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo err:
'MsgBox isPayment

    If isPayment = True Then
        If rs.State = adStateOpen Then rs.Close
      
      'sql = "SELECT tblPayment.TransactionID, tblPayment.RoomRate, tblPayment.NoDays, Last(tblPayment.AmountDue) AS LastOfAmountDue, Sum(tblPayment.cPaid) AS SumOfcPaid, [LastOfAmountDue]-[SumOfcPaid] AS Expr1" & _
            "From tblPayment GROUP BY tblPayment.TransactionID, tblPayment.RoomRate, tblPayment.NoDays, [LastOfAmountDue]-[SumOfcPaid];"
                      
        rs.Open "SELECT * FROM tblPayment WHERE TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        rs.AddNew

        rs.Fields("TransactionID").Value = transID
        rs.Fields("RoomRate").Value = Me.txtRoomRate.Text
        rs.Fields("AmountDue").Value = Me.txtAmountDue.Text
        rs.Fields("RoomNo").Value = isRoomNo
        rs.Fields("PType").Value = Me.cboType.Text
        rs.Fields("Balance").Value = Me.txtBalance.Text
        rs.Fields("NoDays").Value = Me.cboDays.Text
        rs.Fields("cPaid").Value = CCur(Me.txtPayment.Text)
        rs.Fields("pDate").Value = FormatDateTime(Now, vbShortDate)
        rs.Update
        Me.cmdSave.Enabled = False
        Me.cmdClose.Enabled = True
    Else
        Me.cmdSave.Enabled = False
    End If
        chkLock
        sCollection
    Exit Sub
    
    
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Load()

cbo_days.Init Me.cboDays

reposition

End Sub
Private Sub reposition()
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2) + (5880 / 2)
End Sub

Private Sub getdatetime()

End Sub

Public Sub getAmount()
End Sub


Public Sub chkSummary()
If CCur(Me.txtBalance.Text) <= 0 Then
    Me.lblPaymentSummary.Caption = "Paid"
ElseIf CCur(Me.txtBalance.Text) > 0 Then
    Me.lblPaymentSummary.Caption = "Unpaid"
    Me.txtPayment.Enabled = True
End If
End Sub

Private Sub chkisPayment()
    If isPayment = False Then
        Me.txtBalance.Text = FormatNumber(Me.txtAmountDue.Text, 2)
    End If
End Sub

Private Sub txtAmountDue_Change()
On Error GoTo err:
    If isPayment = False Then
        Me.txtBalance.Text = FormatNumber(CDbl(Me.txtAmountDue.Text) - CDbl(Me.txtPayment.Text), 2)
    'Else
        'cBalance
    End If
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub txtBalance_Change()
chkSummary
End Sub

Private Sub txtPayment_GotFocus()
    HLTxt txtPayment
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
           cmdSave_Click
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPayment_LostFocus()
    Me.txtPayment.Text = FormatNumber(txtPayment.Text, 2)
End Sub

Private Sub SaveBills()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM tblPayment WHERE TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
    'MsgBox Me.txtBalance.Text
    rs.Fields("Balance").Value = Me.txtBalance.Text
    rs.Fields("AmountDue").Value = Me.txtAmountDue.Text
    rs.Fields("NoDays").Value = Me.cboDays.Text
    rs.Fields("PType").Value = Me.cboType.Text
    rs.Fields("RoomRate").Value = Me.txtRoomRate.Text
    rs.Fields("cPaid").Value = Me.txtPayment.Text
    rs.Update
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub LoadDateTime()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblCustomerInfo where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        Me.lbldate = FormatDateTime(rs.Fields("occu_date").Value, vbLongDate)
        Me.lblTime = FormatDateTime(rs.Fields("occu_time").Value, vbLongTime)
        dtin = rs.Fields("occu_date").Value
        stime = Format(rs.Fields("occu_time").Value, "HHMMSS")
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub chkLock()
    If CCur(Me.txtBalance.Text) <= 0 Then
        Me.txtPayment.Locked = True
    ElseIf CCur(Me.txtBalance.Text) > 0 Then
        Me.txtPayment.Locked = False
    End If
End Sub
Private Sub sCollection()
On Error GoTo err:
      If rs.State = adStateOpen Then rs.Close
      
        rs.Open "SELECT * FROM tblCollection", cn, adOpenKeyset, adLockPessimistic
        rs.AddNew
        
        rs.Fields("TransID").Value = transID
        rs.Fields("Amount").Value = CCur(Me.txtPayment.Text) '+ CCur(Me.lblPaid.Caption)
        rs.Fields("dDate").Value = FormatDateTime(Now, vbShortDate)
        rs.Fields("dTime").Value = FormatDateTime(Now, vbLongTime)
        rs.Fields("pType").Value = Me.cboType.Text
        rs.Update
    Exit Sub
    
    
err:
    MsgBox err.Description, vbCritical
End Sub
