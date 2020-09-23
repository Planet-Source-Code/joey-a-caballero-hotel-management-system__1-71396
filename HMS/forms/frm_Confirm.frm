VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Confirm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotel Management System [-BPH-]"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin HMS.b8SContainer b8SContainer1 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      BorderColor     =   14215660
      Begin VB.ComboBox cboDep 
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
         ItemData        =   "frm_Confirm.frx":0000
         Left            =   2400
         List            =   "frm_Confirm.frx":000A
         TabIndex        =   5
         Text            =   "Yes"
         Top             =   930
         Width           =   975
      End
      Begin HMS.b8Line b8Line1 
         Height          =   60
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   106
      End
      Begin lvButton.lvButtons_H cmdOk 
         Height          =   405
         Left            =   3720
         TabIndex        =   0
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         Caption         =   "&OK"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keycard Deposit :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer's Data successfully saved. Thank you"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   4035
      End
      Begin VB.Image Image1 
         Height          =   600
         Left            =   120
         Picture         =   "frm_Confirm.frx":0017
         Stretch         =   -1  'True
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frm_Confirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sTrans As Integer
Private Sub cmdOk_Click()
If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblCustomerInfo where TransactionID =" & sTrans, cn, adOpenKeyset, adLockPessimistic
    rs.Fields("deposit").Value = Me.cboDep.Text
    rs.Update
    Unload frm_checkin
    Unload Me
End Sub

Private Sub Form_Load()
If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblCustomerInfo", cn, adOpenKeyset, adLockPessimistic
    rs.MoveLast
    sTrans = rs.Fields("TransactionID").Value
End Sub
