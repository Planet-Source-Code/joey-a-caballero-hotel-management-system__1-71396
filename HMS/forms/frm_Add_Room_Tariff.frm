VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Add_Room_Tariff 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Room Tariff Entry Section"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvlList 
      Height          =   3735
      Left            =   360
      TabIndex        =   12
      Top             =   1800
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Services"
         Object.Width           =   5557
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entry Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   4815
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   2
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox txtMeals 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2950
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   3240
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Tariff"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
   End
   Begin HMS.b8SContainer b8SContainer1 
      Height          =   810
      Left            =   0
      TabIndex        =   9
      Top             =   5640
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1429
      BorderColor     =   14737632
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   4080
         TabIndex        =   6
         Top             =   210
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
         TabIndex        =   5
         Top             =   210
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
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   405
         Left            =   2175
         TabIndex        =   3
         Top             =   210
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
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   405
         Left            =   1215
         TabIndex        =   4
         Top             =   210
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
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdAdd 
         Default         =   -1  'True
         Height          =   405
         Left            =   255
         TabIndex        =   0
         Top             =   210
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Room Tariff Item"
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
      TabIndex        =   10
      Top             =   120
      Width           =   3120
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frm_Add_Room_Tariff.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7185
   End
End
Attribute VB_Name = "frm_Add_Room_Tariff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim isEditM As Boolean
Private Sub cmdAdd_Click()
    unlockDetails
    isEditM = False
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
    
    cn.Execute "DELETE FROM tblRoomRate WHERE ID = " & Me.lvlList.SelectedItem.Text
    'Me.lvwUser.SetFocus
    'Me.lvwUser.ListItems.Item(1).Selected = True
    'lvwUser_ItemClick
    Call loadMeals
    MsgBox "Record deleted.", vbInformation

    Exit Sub
    
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub cmdEdit_Click()
    isEditM = True
    Me.txtMeals.Enabled = True
    Me.txtPrice.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = True
    cmdAdd.Enabled = False
End Sub

Private Sub cmdSave_Click()
    With rs
        If isEditM = False Then         'add record to database
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM tblRoomRate WHERE ID = 0;", cn, adOpenKeyset, adLockOptimistic
            .AddNew
            .Fields("Room_Type") = Me.txtMeals.Text
            .Fields("Room_Rate") = Me.txtPrice.Text
            .Update
            MsgBox "New record saved.", vbInformation
        ElseIf isEditM = True Then    'update current record
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM tblRoomRate WHERE ID = " & Me.lvlList.SelectedItem.Text, cn, adOpenKeyset, adLockOptimistic
            .Fields("Room_Type") = Me.txtMeals.Text
            .Fields("Room_Rate") = Me.txtPrice.Text
            .Update
            MsgBox "Record updated.", vbInformation
        End If
        loadMeals
        lockDetails
    End With
End Sub

Private Sub Form_Load()
    loadMeals
    lvlList.Enabled = True
End Sub

Private Sub loadMeals()
On Error GoTo err:
     If rs.State = adStateOpen Then rs.Close
     rs.Open "Select * from tblRoomRate", cn, adOpenKeyset, adLockPessimistic
     Me.lvlList.ListItems.Clear
     Do While rs.EOF = False
        Me.lvlList.ListItems.Add , , rs.Fields("ID").Value
        Me.lvlList.ListItems(Me.lvlList.ListItems.Count).SubItems(1) = rs.Fields("Room_Type").Value
        Me.lvlList.ListItems(Me.lvlList.ListItems.Count).SubItems(2) = FormatNumber(rs.Fields("Room_Rate").Value, 2)
     rs.MoveNext
     Loop
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub lockDetails()
    Me.txtMeals.Text = ""
    Me.txtMeals.Enabled = False
    Me.txtPrice.Text = "0.00"
    Me.txtPrice.Enabled = False
    
    cmdSave.Enabled = False
    cmdEdit.Enabled = False
    cmdAdd.Enabled = True
   ' cmdCancel.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub unlockDetails()
    Me.txtMeals.Text = ""
    Me.txtMeals.Enabled = True
    Me.txtPrice.Text = "0.00"
    Me.txtPrice.Enabled = True
    
    cmdSave.Enabled = True
    cmdEdit.Enabled = True
    cmdAdd.Enabled = False
    'cmdCancel.Enabled = True
    cmdDelete.Enabled = True
End Sub

Private Sub lvlList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdEdit.Enabled = True
Me.txtMeals.Text = Me.lvlList.ListItems(Me.lvlList.SelectedItem.Index).SubItems(1)
Me.txtPrice.Text = FormatNumber(Me.lvlList.ListItems(Me.lvlList.SelectedItem.Index).SubItems(2), 2)
End Sub

Private Sub txtMeals_GotFocus()
    HLTxt txtMeals
End Sub

Private Sub txtMeals_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtPrice_GotFocus()
    HLTxt txtPrice
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
