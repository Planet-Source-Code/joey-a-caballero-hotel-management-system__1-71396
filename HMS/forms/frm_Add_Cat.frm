VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_Add_Cat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meals Category Entry Section"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvlList 
      Height          =   2655
      Left            =   360
      TabIndex        =   10
      Top             =   1560
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   4683
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Meals Category"
         Object.Width           =   5998
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
      Height          =   3735
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   3975
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
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
   End
   Begin HMS.b8SContainer b8SContainer1 
      Height          =   930
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1640
      BorderColor     =   14737632
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   3240
         TabIndex        =   5
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   2520
         TabIndex        =   4
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   1800
         TabIndex        =   2
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   1095
         TabIndex        =   3
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   360
         TabIndex        =   0
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
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
      Caption         =   "Add Meals Category"
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
      TabIndex        =   9
      Top             =   120
      Width           =   2850
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frm_Add_Cat.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4545
   End
End
Attribute VB_Name = "frm_Add_Cat"
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
    
    cn.Execute "DELETE FROM tblMealCat WHERE ID = " & Me.lvlList.SelectedItem.Text
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
    
    cmdSave.Enabled = True
    cmdEdit.Enabled = True
    cmdAdd.Enabled = False
End Sub

Private Sub cmdSave_Click()
    With rs
        If isEditM = False Then         'add record to database
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM tblMealCat WHERE ID = 0;", cn, adOpenKeyset, adLockOptimistic
            .AddNew
            .Fields("MCat") = Me.txtMeals.Text
            .Update
            MsgBox "New record saved.", vbInformation
        ElseIf isEditM = True Then    'update current record
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM tblMealCat WHERE ID = " & Me.lvlList.SelectedItem.Text, cn, adOpenKeyset, adLockOptimistic
            .Fields("MCat") = Me.txtMeals.Text
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
     rs.Open "Select * from tblMealCat", cn, adOpenKeyset, adLockPessimistic
     Me.lvlList.ListItems.Clear
     Do While rs.EOF = False
        Me.lvlList.ListItems.Add , , rs.Fields("ID").Value
        Me.lvlList.ListItems(Me.lvlList.ListItems.Count).SubItems(1) = rs.Fields("MCat").Value
     rs.MoveNext
     Loop
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub lockDetails()
    Me.txtMeals.Text = ""
    Me.txtMeals.Enabled = False
    
    cmdSave.Enabled = False
    cmdEdit.Enabled = False
    cmdAdd.Enabled = True
   ' cmdCancel.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub unlockDetails()
    Me.txtMeals.Text = ""
    Me.txtMeals.Enabled = True
    
    cmdSave.Enabled = True
    cmdEdit.Enabled = True
    cmdAdd.Enabled = False
    'cmdCancel.Enabled = True
    cmdDelete.Enabled = True
End Sub

Private Sub lvlList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.cmdEdit.Enabled = True
Me.txtMeals.Text = Me.lvlList.ListItems(Me.lvlList.SelectedItem.Index).SubItems(1)
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
