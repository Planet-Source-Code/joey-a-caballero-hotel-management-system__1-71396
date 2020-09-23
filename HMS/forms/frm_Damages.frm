VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Damages 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Damages"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   240
      ScaleHeight     =   0
      ScaleWidth      =   7035
      TabIndex        =   1
      Top             =   5400
      Width           =   7095
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Price"
         Object.Width           =   1649
      EndProperty
   End
   Begin MSComctlLib.ListView lvwlistAdd 
      Height          =   2175
      Left            =   3120
      TabIndex        =   8
      Top             =   3000
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Menu"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Price"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   1676
      EndProperty
   End
   Begin MSComctlLib.ListView lvwListSel 
      Height          =   1935
      Left            =   3120
      TabIndex        =   9
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Menu"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Price"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   1676
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   5640
      Width           =   1095
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
      Left            =   2400
      TabIndex        =   7
      Top             =   660
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
      Left            =   2400
      TabIndex        =   6
      Top             =   405
      Width           =   45
   End
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
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frm_Damages.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
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
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   840
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
      Left            =   1200
      TabIndex        =   3
      Top             =   405
      Width           =   810
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
      Left            =   1200
      TabIndex        =   2
      Top             =   660
      Width           =   1035
   End
End
Attribute VB_Name = "frm_Damages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdAdd_Click()
Dim cntA As Integer
    For x = 1 To Me.lvwList.ListItems.Count
        If Me.lvwList.ListItems(x).Checked = True Then
            Me.lvwListSel.ListItems.Add , , Me.lvwList.ListItems(x).Text
            Me.lvwListSel.ListItems(Me.lvwListSel.ListItems.Count).SubItems(1) = Me.lvwList.ListItems(x).SubItems(1)
            Me.lvwListSel.ListItems(Me.lvwListSel.ListItems.Count).SubItems(2) = FormatDateTime(Now, vbShortDate)
            Me.lvwListSel.ListItems(Me.lvwListSel.ListItems.Count).SubItems(3) = FormatDateTime(Now, vbShortTime)
            'rs.Update
            cntA = cntA + 1
        End If
    Next x
        If cntA = 0 Then
            MsgBox "Nothing to add", vbInformation
        Else
            MsgBox "Selected Item Added. Thank You!", vbInformation, "Information"
            Me.cmdSave.Enabled = True
        End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim x As Integer
On Error GoTo err:
    If MsgBox("Saving selected item cannot be undo." & vbCrLf & vbCrLf & "Are you sure?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblAddOnDamage", cn, adOpenKeyset, adLockPessimistic
            For x = 1 To Me.lvwListSel.ListItems.Count
                    rs.AddNew
                    rs.Fields("TransID").Value = transID
                    rs.Fields("Damage").Value = Me.lvwListSel.ListItems(x).Text
                    rs.Fields("Cost").Value = Me.lvwListSel.ListItems(x).SubItems(1)
                    rs.Fields("dDate").Value = FormatDateTime(Now, vbShortDate)
                    rs.Fields("dTime").Value = FormatDateTime(Now, vbShortTime)
                    rs.Update
            Next x
        MsgBox "Selected Item Added. Thank You!", vbInformation, "Information"
        LoadItem
        Me.cmdSave.Enabled = False
    End If
    Me.lvwListSel.ListItems.Clear
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Load()
    loadDamages
    LoadItem
End Sub

Private Sub loadDamages()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblDamages", cn, adOpenKeyset, adLockPessimistic
    Do While rs.EOF = False
        Me.lvwList.ListItems.Add , , rs.Fields("dDamages").Value
        Me.lvwList.ListItems(Me.lvwList.ListItems.Count).SubItems(1) = FormatNumber(rs.Fields("dPrice").Value, 2)
        rs.MoveNext
    Loop
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub LoadItem()
On Error GoTo err:
    Dim x As Integer
    Dim dblTotal As Double
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblAddOnDamage where TransID = " & transID, cn, adOpenKeyset, adLockPessimistic
    Me.lvwlistAdd.ListItems.Clear
    Do While rs.EOF = False
        Me.lvwlistAdd.ListItems.Add , , rs.Fields("Damage").Value
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(1) = FormatNumber(rs.Fields("Cost").Value, 2)
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(2) = rs.Fields("dDate").Value
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(3) = rs.Fields("dTime").Value
        rs.MoveNext
    Loop
    
    If Me.lvwlistAdd.ListItems.Count > 0 Then
        For x = 1 To Me.lvwlistAdd.ListItems.Count
            dblTotal = dblTotal + CDbl(Me.lvwlistAdd.ListItems(x).SubItems(1))
        Next x
        Me.lvwlistAdd.ListItems.Add , , ""
        Me.lvwlistAdd.ListItems.Add , , "---------------------------------------"
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(1) = "-----------------"
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(2) = "-----------------"
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(3) = "-----------------"
        Me.lvwlistAdd.ListItems.Add , , "Total Cost"
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(1) = FormatNumber(dblTotal, 2)
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(2) = FormatDateTime(Now, vbShortDate)
        Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(3) = FormatDateTime(Now, vbLongTime)
    End If
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim x As Integer
    On Error GoTo err:
    If Me.lvwlistAdd.ListItems.Count <> 0 Then
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblPayment where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        rs.Fields("Damages").Value = Me.lvwlistAdd.ListItems(Me.lvwlistAdd.ListItems.Count).SubItems(1)
        rs.Update
    End If
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub lvwListSel_DblClick()
If Me.lvwListSel.ListItems.Count <> 0 Then
    Me.lvwListSel.ListItems.Remove (Me.lvwListSel.SelectedItem.Index)
End If

If Me.lvwListSel.ListItems.Count = 0 Then
    Me.cmdSave.Enabled = False
End If
End Sub
