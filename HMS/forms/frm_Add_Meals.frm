VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Add_Meals 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meals Entry Section"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvlList 
      Height          =   4575
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   8070
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Meals"
         Object.Width           =   5557
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
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
      Height          =   5895
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   6615
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
         Left            =   5040
         TabIndex        =   3
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
         Left            =   1980
         TabIndex        =   2
         Top             =   720
         Width           =   2950
      End
      Begin VB.ComboBox cboCat 
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
         Width           =   1695
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
         Left            =   5040
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Meals"
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
         Left            =   1980
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin HMS.b8SContainer b8SContainer1 
      Height          =   810
      Left            =   0
      TabIndex        =   11
      Top             =   6480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1429
      BorderColor     =   14737632
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   5760
         TabIndex        =   7
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
         Left            =   4815
         TabIndex        =   6
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
         Left            =   3855
         TabIndex        =   4
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
         Left            =   2895
         TabIndex        =   5
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
         Left            =   1935
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
      Caption         =   "Add Meals"
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
      TabIndex        =   12
      Top             =   120
      Width           =   1470
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frm_Add_Meals.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7185
   End
End
Attribute VB_Name = "frm_Add_Meals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim isEditM As Boolean
Dim cboAutoCat As New CAutoCompleteComboBox

Private Sub cboCat_GotFocus()
    HLTxt cboCat
End Sub

Private Sub cboCat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

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
    
    cn.Execute "DELETE FROM tblMeals WHERE ID = " & Me.lvlList.SelectedItem.Text
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
            .Open "SELECT * FROM tblMeals WHERE ID = 0;", cn, adOpenKeyset, adLockOptimistic
            .AddNew
            .Fields("MMeals") = Me.txtMeals.Text
            .Fields("MPrice") = Me.txtPrice.Text
            .Fields("MCat") = Me.cboCat.Text
            .Update
            MsgBox "New record saved.", vbInformation
        ElseIf isEditM = True Then    'update current record
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM tblMeals WHERE ID = " & Me.lvlList.SelectedItem.Text, cn, adOpenKeyset, adLockOptimistic
            .Fields("MMeals") = Me.txtMeals.Text
            .Fields("MPrice") = Me.txtPrice.Text
            .Fields("MCat") = Me.cboCat.Text
            .Update
            MsgBox "Record updated.", vbInformation
        End If
        loadMeals
        lockDetails
    End With
End Sub

Private Sub Form_Load()
    loadCat
    loadMeals
    lvlList.Enabled = True
    cboAutoCat.Init Me.cboCat
End Sub

Private Sub loadCat()
On Error GoTo err:
     If rs.State = adStateOpen Then rs.Close
     rs.Open "Select * from tblMealCat", cn, adOpenKeyset, adLockPessimistic
     Do While rs.EOF = False
        Me.cboCat.AddItem rs.Fields("MCat").Value
     rs.MoveNext
     Loop
        
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub loadMeals()
On Error GoTo err:
     If rs.State = adStateOpen Then rs.Close
     rs.Open "Select * from tblMeals", cn, adOpenKeyset, adLockPessimistic
     Me.lvlList.ListItems.Clear
     Do While rs.EOF = False
        Me.lvlList.ListItems.Add , , rs.Fields("ID").Value
        Me.lvlList.ListItems(Me.lvlList.ListItems.Count).SubItems(1) = rs.Fields("MCat").Value
        Me.lvlList.ListItems(Me.lvlList.ListItems.Count).SubItems(2) = rs.Fields("MMeals").Value
        Me.lvlList.ListItems(Me.lvlList.ListItems.Count).SubItems(3) = FormatNumber(rs.Fields("MPrice").Value, 2)
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
Me.cboCat.Text = Me.lvlList.ListItems(Me.lvlList.SelectedItem.Index).SubItems(1)
Me.txtMeals.Text = Me.lvlList.ListItems(Me.lvlList.SelectedItem.Index).SubItems(2)
Me.txtPrice.Text = FormatNumber(Me.lvlList.ListItems(Me.lvlList.SelectedItem.Index).SubItems(3), 2)
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
