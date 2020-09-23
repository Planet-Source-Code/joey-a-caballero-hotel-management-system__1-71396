VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Payment 
   BorderStyle     =   0  'None
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvlList 
      Height          =   735
      Left            =   600
      TabIndex        =   31
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
            Picture         =   "frm_Payment.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Payment.frx":5284
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bgMain 
      Height          =   6195
      Left            =   0
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   347
      TabIndex        =   9
      Top             =   0
      Width           =   5265
      Begin VB.CheckBox chkNonTax 
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
         TabIndex        =   32
         Top             =   3000
         Width           =   4815
      End
      Begin MSComctlLib.ListView lvlReport 
         Height          =   1455
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2566
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer's Name"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "No. of Days"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Payments"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Discount"
            Object.Width           =   1720
         EndProperty
      End
      Begin VB.TextBox txtDiscount 
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
         Left            =   3720
         TabIndex        =   29
         Text            =   "0.0"
         Top             =   2160
         Width           =   1215
      End
      Begin HMS.b8Line b8Line2 
         Height          =   60
         Left            =   0
         TabIndex        =   23
         Top             =   3360
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
         ItemData        =   "frm_Payment.frx":A198
         Left            =   3720
         List            =   "frm_Payment.frx":A1A5
         TabIndex        =   7
         Text            =   "Cash"
         Top             =   2520
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
         Left            =   1440
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1800
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
         Left            =   1440
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox cboDays 
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
         ItemData        =   "frm_Payment.frx":A1BA
         Left            =   3720
         List            =   "frm_Payment.frx":A218
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Text            =   "0"
         Top             =   1440
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
               Picture         =   "frm_Payment.frx":A28B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HMS.b8SContainer pbBGButton 
         Height          =   870
         Left            =   -120
         TabIndex        =   10
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1535
         BorderColor     =   14215660
         Begin VB.Label lblPaid 
            Caption         =   "0.00"
            Height          =   255
            Left            =   3840
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblcBalance 
            Caption         =   "0.00"
            Height          =   255
            Left            =   3840
            TabIndex        =   26
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   120
            Picture         =   "frm_Payment.frx":A825
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   240
            Width           =   45
         End
      End
      Begin HMS.b8SContainer b8SContainer1 
         Height          =   870
         Left            =   0
         TabIndex        =   15
         Top             =   5280
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1535
         BorderColor     =   14215660
         Begin HMS.b8Line b8Line1 
            Height          =   60
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   106
         End
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   405
            Left            =   3360
            TabIndex        =   8
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount:"
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
         Left            =   2760
         TabIndex        =   28
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label lblPaymentSummary 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unpaid"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   65.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   4935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Summary:"
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
         TabIndex        =   24
         Top             =   3480
         Width           =   1890
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
         TabIndex        =   22
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1275
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Days:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2760
         TabIndex        =   18
         Top             =   1440
         Width           =   810
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
         Left            =   2760
         TabIndex        =   17
         Top             =   2520
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rsClear As New ADODB.Recordset
Dim roomrate As Currency
Dim cbo_days As New CAutoCompleteComboBox
Dim paid As Currency
Dim bal As Currency
Dim dtin, dtnow As Date
Dim stime As Double
Dim lTime As Double
Dim addDamage As Currency
Dim addMeals As Currency
Dim addServices As Currency
Dim x As Currency
Dim zRated As Currency

Private Sub b8Title_GotFocus()
 mdiMain.Enabled = False
End Sub

Private Sub bgMain_GotFocus()
 mdiMain.Enabled = False
End Sub

Private Sub cboDays_Change()
On Error Resume Next
     If isPayment = True Then
            Me.txtAmountDue.Text = FormatNumber(CDbl(cboDays.Text) * CDbl(Me.txtRoomRate.Text) + addDamage + addMeals + addServices, 2)
            If Me.chkNonTax.Value = 1 Then
                zRated = FormatNumber((CCur(Me.txtRoomRate.Text) * CDbl(Me.cboDays.Text) / 1.12) + addDamage + addMeals + addServices, 2)
                 Me.txtBalance.Text = FormatNumber(CCur(zRated - CCur(Me.lblPaid.Caption)), 2)
            Else
                 Me.txtBalance.Text = FormatNumber(CCur(Me.txtAmountDue.Text - CCur(Me.lblPaid.Caption)), 2)
            End If
           
        ''''''
        ''''''
        'lblcBalance.Caption = Me.txtBalance.Text
        Me.cmdSave.Enabled = True
        Me.cmdClose.Enabled = False
        chkSummary
    Else
        
            On Error Resume Next
            Me.txtAmountDue.Text = FormatNumber(CDbl(cboDays.Text) * CDbl(Me.txtRoomRate.Text), 2)
    End If
End Sub

Private Sub cboDays_Click()
'    If isPayment = True Then
'        Me.txtAmountDue.Text = FormatNumber(CDbl(cboDays.Text) * CDbl(Me.txtRoomRate.Text) + addDamage + addMeals + addServices, 2)
'        Me.txtBalance.Text = FormatNumber(CCur(Me.txtAmountDue.Text - CCur(Me.lblPaid.Caption)), 2)
        ''''''
        ''''''
        'lblcBalance.Caption = Me.txtBalance.Text
'        Me.cmdSave.Enabled = True
'        Me.cmdClose.Enabled = False
'        chkSummary
'    Else
'        Me.txtAmountDue.Text = FormatNumber(CDbl(cboDays.Text) * CDbl(Me.txtRoomRate.Text), 2)
'    End If
End Sub
''''''''''''''
'''''''''''''''
''''''''''''''''
'''''''''''''''
''''''''''''''

Private Sub chkNonTax_Click()
    If Me.chkNonTax.Value = 1 Then
        zRated = FormatNumber((CCur(Me.txtRoomRate.Text) * CDbl(Me.cboDays.Text) / 1.12) + addDamage + addMeals + addServices, 2)
    End If
    
    txtDiscount_Change
'MsgBox Me.txtAmountDue.Text
End Sub

Private Sub cmdClose_Click()
    'MsgBox ischkout
    If isPayment = False Then
    'MsgBox isPayment
        SaveBills
    End If
    
    If ischkout = True Then
        If Me.lblPaymentSummary.Caption = "Paid" Then
            mdiMain.pcheckout
        End If
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
        'Me.txtRoomRate.Text = rs.Fields("RoomRate").Value
        rs.Fields("TransactionID").Value = transID
        rs.Fields("RoomRate").Value = Me.txtRoomRate.Text
        rs.Fields("AmountDue").Value = Me.txtAmountDue.Text
        rs.Fields("RoomNo").Value = isRoomNo
        rs.Fields("PType").Value = Me.cboType.Text
        rs.Fields("Balance").Value = Me.txtBalance.Text
        rs.Fields("NoDays").Value = Me.cboDays.Text
        rs.Fields("cPaid").Value = CCur(Me.txtPayment.Text) '+ CCur(Me.lblPaid.Caption)
        rs.Fields("pDate").Value = FormatDateTime(Now, vbShortDate)
        rs.Update
        Me.cmdSave.Enabled = False
        Me.cmdClose.Enabled = True
    Else
        Me.cmdSave.Enabled = False
    End If
        chkLock
        sDiscount
        'Me.lvlReport.ListItems.Clear
        loadSummary
        sCollection
    Exit Sub
    
    
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Activate()
    chkLock
End Sub

''''''''''''''
'''''''''''''''
''''''''''''''''
'''''''''''''''
''''''''''''''
Private Sub Form_Load()
'SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
'SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_ALPHA
'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
cbo_days.Init Me.cboDays

reposition
'MsgBox isPayment

If isRchkin = True Then
'MsgBox "here:"
    'ConfirmChkin
Else
    If isPayment = True Then
        LoadAdditional
        LoadDateTime
        loadData
        getDiscount
        getNoDays
        chkNonTax_Click
        loadSummary
        Me.cmdClose.Enabled = True
        Me.cmdSave.Enabled = False
    Else
        'getNoDays
        getdatetime
        getAmount
        chkNonTax_Click
        lTime = Format(Time, "HHMMSS")
        chkisPayment
    End If

    chkSummary
    numDays = Me.cboDays.Text
    If CCur(Me.txtBalance.Text) <= 0 Then
        Me.cmdClose = True
    End If
End If
End Sub
Private Sub reposition()
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2) + (5880 / 2)
End Sub

Private Sub getdatetime()
    Me.lbldate = FormatDateTime(Now, vbLongDate)
    Me.lblTime = Time
End Sub

Public Sub getAmount()
On Error GoTo err:
    'MsgBox isPayment & " " & ischkout
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM tblRoomRate WHERE Room_Type = '" & roomtype & "'", cn, adOpenKeyset, adLockPessimistic
        Me.txtRoomRate.Text = FormatNumber(rs.Fields("Room_Rate").Value, 2)
         If Me.chkNonTax.Value = 1 Then
            zRated = FormatNumber((CDbl(cboDays.Text) * CDbl(Me.txtRoomRate.Text) / 1.12) + addDamage + addMeals + addServices, 2)
            Me.txtAmountDue.Text = zRated
        Else
            Me.txtAmountDue.Text = FormatNumber(CDbl(cboDays.Text) * CDbl(Me.txtRoomRate.Text) + addDamage + addMeals + addServices, 2)
        End If
        
Exit Sub
err:
    MsgBox err.Description, vbCritical
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 mdiMain.Enabled = True
End Sub

Private Sub pbBGButton_GotFocus()
 mdiMain.Enabled = False
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

Private Sub txtDiscount_Change()
    'x = CDbl(Me.Text1.Text / 100) * CDbl(Me.Text2.Text)
    'z = CDbl(Me.Text1.Text) - x
    'MsgBox z
    
On Error GoTo err:
    If Me.txtDiscount.Text = vbNullString Then
        Me.txtDiscount.Text = FormatNumber(0, 1)
    End If
    
        calDiscount
        If Me.chkNonTax.Value = 1 Then
            zRated = FormatNumber((CCur(Me.txtRoomRate.Text) * CDbl(Me.cboDays.Text) / 1.12) + addDamage + addMeals + addServices, 2)
            Me.txtBalance.Text = FormatNumber(CCur(zRated) - CCur(txtPayment.Text) - CCur(Me.lblPaid.Caption) - x, 2)
        Else
            Me.txtBalance.Text = FormatNumber(CCur(Me.txtAmountDue.Text) - CCur(txtPayment.Text) - CCur(Me.lblPaid.Caption) - x, 2)
        End If
        
        'Me.txtBalance.Text = FormatNumber(CCur(Me.txtAmountDue.Text) - CCur(txtPayment.Text) - CCur(Me.lblPaid.Caption) - CCur(Me.txtDiscount.Text), 2)
    Exit Sub
err:
Me.txtDiscount.Text = FormatNumber(0, 1)
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            Me.txtPayment.SetFocus
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtDiscount_LostFocus()
    Me.txtDiscount.Text = FormatNumber(txtDiscount.Text, 1)
End Sub

Private Sub txtPayment_Change()
On Error GoTo err:
    If Me.txtPayment.Text = vbNullString Then
        Me.txtPayment.Text = FormatNumber(0, 2)
    End If
    If isPayment = False Then
        If Me.chkNonTax.Value = 1 Then
            zRated = FormatNumber((CCur(Me.txtRoomRate.Text) * CDbl(Me.cboDays.Text) / 1.12) + addDamage + addMeals + addServices, 2)
            Me.txtBalance.Text = FormatNumber(CCur(zRated) - CCur(txtPayment.Text) - x, 2)
        Else
            Me.txtBalance.Text = FormatNumber(CCur(Me.txtAmountDue.Text) - CCur(Me.txtPayment.Text) - x, 2)
        End If
    Else
        If Me.chkNonTax.Value = 1 Then
            zRated = FormatNumber((CCur(Me.txtRoomRate.Text) * CDbl(Me.cboDays.Text) / 1.12) + addDamage + addMeals + addServices, 2)
            Me.txtBalance.Text = FormatNumber(CCur(zRated) - CCur(txtPayment.Text) - CCur(Me.lblPaid.Caption) - x, 2)
        Else
            Me.txtBalance.Text = FormatNumber(CCur(Me.txtAmountDue.Text) - CCur(txtPayment.Text) - CCur(Me.lblPaid.Caption) - x, 2)
        End If
    End If
    'Me.lblcBalance.Caption = Me.txtBalance.Text
    Me.cmdSave.Enabled = True
    Exit Sub
err:
Me.txtPayment.Text = FormatNumber(0, 2)
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

Private Sub loadData()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    Dim sql As String
    
    rs.Open "SELECT tblPayment.RoomRate, Max(tblPayment.NoDays) AS MaxOfNoDays, Sum(tblPayment.cPaid) AS SumOfcPaid, Max(tblPayment.AmountDue) AS MaxOfAmountDue, Max(tblPayment.AmountDue)-Sum(tblPayment.cPaid) AS Expr1, tblPayment.TransactionID " & _
"From tblPayment Where (((tblPayment.TransactionID) = " & transID & "))" & _
" GROUP BY tblPayment.RoomRate, tblPayment.TransactionID;", cn, adOpenKeyset, adLockPessimistic

    'MsgBox rs.Fields("SumOfcPaid").Value
    paid = rs.Fields("SumOfcPaid").Value
    lblPaid.Caption = FormatNumber(paid, 2)
    Me.lblcBalance.Caption = FormatNumber(rs.Fields("Expr1").Value, 2)
    'MsgBox Me.lblcBalance.Caption
    Me.txtBalance.Text = FormatNumber(Me.lblcBalance.Caption, 2)
    Me.txtRoomRate.Text = FormatNumber(rs.Fields("RoomRate").Value, 2)
    Me.cboDays.Text = rs.Fields("MaxOfNoDays").Value
    
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

'Private Sub cBalance()
'On error goto err:
'Dim sql As String

'    If rs.State = adStateOpen Then rs.Close
'    sql = "Select * from tblPayment where TransactionID=" & transID & ";"
'    rs.Open sql, cn, adOpenKeyset, adLockPessimistic
    'bal = rs.Fields("Balance").Value
    'paid = rs.Fields("cPaid").Value
    'lblcBalance.Caption = bal
    'lblPaid.Caption = paid
    'Me.txtBalance.Text = FormatNumber(Me.txtAmountDue.Text - lblcBalance.Caption)
 '   Exit Sub
'err:
'    MsgBox err.Description, vbCritical
'End Sub

Private Sub txtRoomRate_GotFocus()
    mdiMain.Enabled = False
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

Public Sub getNoDays()
    dtnow = FormatDateTime(Now, vbShortDate)
    ' Check-in form 12:00am - 5:00am
    ' Code Ok
    ' For Adjustmet
    If stime >= 100 And stime <= 50000 Then
        If DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 1.5
        ElseIf DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 140000 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 2
        Else
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 1
        End If
    ' Check-in form 5:01am - 7:00am
    ' Code Ok
    ' For Adjustmet
    ElseIf stime >= 50100 And stime <= 70000 Then
        If DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") <= 120100 Then 'and Format(Time, "HHMMSS") <= 140000 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 0.5
        
        ElseIf DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") >= 120100 Then 'and Format(Time, "HHMMSS") <= 140000 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 1
                   
        ElseIf DateDiff("d", dtin, dtnow) > 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 0.5
             
        ElseIf DateDiff("d", dtin, dtnow) > 0 And Format(Time, "HHMMSS") >= 140100 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 1
            
        Else
            Me.cboDays.Text = DateDiff("d", dtin, dtnow)
        End If
    ' Will Follow
    ' Code Ok
    ' For Adjustmet
    ElseIf DateDiff("d", dtin, dtnow) = 0 And stime >= 70000 And stime <= 120000 Then
        If DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") <= 140000 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 0.5
        ElseIf DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") >= 140100 Then
            Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 1
        End If
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
        Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 0.5
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") >= 140100 Then
        Me.cboDays.Text = DateDiff("d", dtin, dtnow) + 1
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") <= 120100 Then
    
        Me.cboDays.Text = DateDiff("d", dtin, dtnow)
    End If
        'gDays = Me.cboDays.Text
End Sub

Private Sub chkLock()
    If CCur(Me.txtBalance.Text) <= 0 Then
        Me.txtPayment.Locked = True
    ElseIf CCur(Me.txtBalance.Text) > 0 Then
        Me.txtPayment.Locked = False
    End If
End Sub

Private Sub ConfirmChkin()
    If rs.State = adStateOpen Then rs.Close
    If rsClear.State = adStateOpen Then rsClear.Close
        
    rsClear.Open "Select * from tblRoomRate where Room_Type = '" & rType & "';", cn, adOpenKeyset, adLockPessimistic
    Me.txtRoomRate.Text = rsClear.Fields("Room_Rate").Value
    
    rs.Open "Select * from tblCustomerInfo where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
    rs.Fields("occu_date").Value = (FormatDateTime(Now, vbShortDate))
    rs.Fields("occu_time").Value = FormatDateTime(Now, vbShortTime)
    rs.Update

    getNoDays
End Sub

Private Sub LoadAdditional()
On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblPayment where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
            addDamage = rs.Fields("Damages").Value
            addMeals = rs.Fields("Meals").Value
            addServices = rs.Fields("Services").Value
    Exit Sub
err:
MsgBox err.Description, vbCritical
        
End Sub

Private Sub getDiscount()
On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblPayment where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        Me.txtDiscount.Text = FormatNumber(rs.Fields("Discount").Value, 1)
        If rs.Fields("isNonTax").Value = "true" Then
            Me.chkNonTax.Value = 1
        Else
            Me.chkNonTax.Value = 0
        End If
        
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub sDiscount()
On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblPayment", cn, adOpenKeyset, adLockPessimistic
        Do While rs.EOF = False
            If rs.Fields("TransactionID").Value = transID Then
               rs.Fields("Discount").Value = Me.txtDiscount.Text
                If Me.chkNonTax.Value = 1 Then
                    rs.Fields("isNonTax").Value = "true"
                Else
                    rs.Fields("isNonTax").Value = "false"
                End If
            End If
        rs.Update
        rs.MoveNext
        Loop
        
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub calDiscount()
    x = (CCur(Me.txtRoomRate.Text) * CDbl(Me.cboDays.Text) / 100) * CCur(Me.txtDiscount.Text)
End Sub

Private Sub loadSummary()
On Error GoTo err:
Dim sql
Dim intCnt As Integer

    If rs.State = adStateOpen Then rs.Close
    
    sql = "SELECT tblPayment.TransactionID, tblPayment.pDate From tblPayment Where (((tblPayment.TransactionID) =" & transID & ")) " & _
          "GROUP BY tblPayment.TransactionID, tblPayment.pDate;"
          
    rs.Open sql, cn, adOpenKeyset, adLockPessimistic
    Me.lvlList.ListItems.Clear
    Do While rs.EOF = False
        Me.lvlList.ListItems.Add , , rs.Fields("pDate").Value
        rs.MoveNext
    Loop
    
    Me.lvlReport.ListItems.Clear
    
    For intCnt = 1 To (Me.lvlList.ListItems.Count)
        If rs.State = adStateOpen Then rs.Close
    

            'sql = "SELECT tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.LastName, tblPayment.NoDays, tblPayment.cPaid, tblPayment.Discount, tblPayment.Meals, tblPayment.Services, tblPayment.Damages, tblPayment.Balance, tblPayment.pDate " & _
            "FROM tblCustomerInfo INNER JOIN tblPayment ON tblCustomerInfo.TransactionID = tblPayment.TransactionID " & _
            "Where (((tblCustomerInfo.TransactionID) =" & transID & ")) " & _
            "GROUP BY tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.LastName, tblPayment.NoDays, tblPayment.cPaid, tblPayment.Discount, tblPayment.Meals, tblPayment.Services, tblPayment.Damages, tblPayment.Balance, tblPayment.pDate;"
            
            sql = "SELECT tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, Sum(tblPayment.cPaid) AS SumOfcPaid, tblPayment.pDate, Sum(tblPayment.Meals) AS SumOfMeals, Sum(tblPayment.Services) AS SumOfServices, Sum(tblPayment.Damages) AS SumOfDamages, tblPayment.NoDays " & _
                  "FROM tblCustomerInfo INNER JOIN tblPayment ON tblCustomerInfo.TransactionID = tblPayment.TransactionID " & _
                  "Where (((tblPayment.pDate) = '" & Me.lvlList.ListItems(intCnt).Text & "') And ((tblPayment.TransactionID) =" & transID & ")) " & _
                  "GROUP BY tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblPayment.pDate, tblPayment.NoDays; "
                  
            rs.Open sql, cn, adOpenKeyset, adLockPessimistic
            
            
            Do While rs.EOF = False
                'If rs.Fields("cPaid").Value = 0 Then
                '    rs.MoveNext
                'Else
                    Me.lvlReport.ListItems.Add , , rs.Fields("pDate").Value
                    Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(1) = rs.Fields("FirstName").Value & " " & rs.Fields("MiddleName").Value & " " & rs.Fields("LastName").Value
                    Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(2) = rs.Fields("NoDays").Value
                    Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(3) = FormatNumber(rs.Fields("SumOfcPaid").Value, 2)
                    Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(4) = Me.txtDiscount.Text
                    'Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(4) = rs.Fields("Discount").Value
                    'Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(4) = rs.Fields("SumOfMeals").Value
                    'Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(5) = rs.Fields("SumOfServices").Value
                    'Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(6) = rs.Fields("SumOfDamages").Value
                    'Me.lvlReport.ListItems(Me.lvlReport.ListItems.Count).SubItems(8) = rs.Fields("Balance").Value
                    rs.MoveNext
                'End If
            Loop
    Next intCnt
Exit Sub
err:
MsgBox err.Description, vbCritical
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
