VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Damage_Report 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Damage Report Manager"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7560
   ControlBox      =   0   'False
   Icon            =   "frm_Damage_Report.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   240
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   720
      Width           =   7065
      Begin MSComctlLib.ImageList icoHeader 
         Left            =   5805
         Top             =   3210
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
               Picture         =   "frm_Damage_Report.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Damage_Report.frx":0B24
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin HMS.b8SContainer b8SConStatus 
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   4620
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   661
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
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
            Left            =   5040
            TabIndex        =   6
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lblRecord 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 Record(s)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   90
            Width           =   990
         End
      End
      Begin HMS.b8SContainer pbBGButton 
         Height          =   630
         Left            =   0
         TabIndex        =   2
         Top             =   345
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1111
         BorderColor     =   14215660
      End
      Begin HMS.b8ChildTitleBar b8Title 
         Height          =   345
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
         BackColor       =   12735512
         Caption         =   "Report Manager"
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
         Left            =   0
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
               Picture         =   "frm_Damage_Report.frx":10BE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   3480
         Left            =   -15
         TabIndex        =   5
         Top             =   915
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6138
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilRecordIco"
         SmallIcons      =   "ilRecordIco"
         ColHdrIcons     =   "icoHeader"
         ForeColor       =   8399906
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Room No."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company  Name"
            Object.Width           =   3307
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer's Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "GR #"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "C/IN Date"
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Damage Item"
            Object.Width           =   3307
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Time"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Damage Cost"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin HMS.b8Container b8cMain 
      Height          =   5940
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10478
      BorderColor     =   12632256
      InsideBorderColor=   14215660
      ShadowColor1    =   14215660
      ShadowColor2    =   14215660
   End
End
Attribute VB_Name = "frm_Damage_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
Dim rs As New ADODB.Recordset
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Sub Form_Load()
    LoadDamageReport
    Rcounter
End Sub

Private Sub Form_Resize()
    ReArrangeControls
End Sub
Public Function Form_Find()

End Function

Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
    Me.WindowState = vbMaximized
End Sub

Private Sub ReArrangeControls()
On Error Resume Next
    
    Me.ScaleMode = vbPixels
    b8cMain.Move Form_LeftMargin - 3, Form_TopMargin - 3, Me.ScaleWidth - (Form_LeftMargin - 3) * 2, Me.ScaleHeight - (Form_TopMargin - 3) * 2
    
    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2
    
    b8Title.Move 0, 0, bgMain.Width
    pbBGButton.Move 0, b8Title.Top + b8Title.Height, bgMain.Width
    listRecord.Move listRecord.Left, pbBGButton.Top + pbBGButton.Height, bgMain.Width - (listRecord.Left * 2)
    listRecord.Height = bgMain.Height - (listRecord.Top + b8SConStatus.Height)
    b8SConStatus.Move -1, bgMain.Height + 1 - b8SConStatus.Height, bgMain.Width + 1
    Me.cmdPrint.Left = Me.Width - Me.cmdPrint.Width - 200


End Sub

Private Sub LoadDamageReport()
On Error GoTo err:
Dim cntD As Integer
Dim curTotal As Currency
Dim intCntX As Integer
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT tblCustomerInfo.Room_no, tblCustomerInfo.grNo, tblCustomerInfo.Company, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.TransactionID, tblCustomerInfo.occu_date, tblAddOnDamage.Damage, tblAddOnDamage.dDate, tblAddOnDamage.dTime, tblAddOnDamage.Cost " & _
            "FROM tblCustomerInfo INNER JOIN tblAddOnDamage ON tblCustomerInfo.TransactionID = tblAddOnDamage.TransID " & _
            "GROUP BY tblCustomerInfo.Room_no, tblCustomerInfo.grNo, tblCustomerInfo.Company, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.TransactionID, tblCustomerInfo.occu_date, tblAddOnDamage.Damage, tblAddOnDamage.dDate, tblAddOnDamage.dTime, tblAddOnDamage.Cost ORDER BY tblCustomerInfo.TransactionID", cn, adOpenKeyset, adLockPessimistic
    cntD = 1
    Do While rs.EOF = False
        'If cntD >= 1 Then
        If Me.listRecord.ListItems.Count <> 0 Then
            If Me.listRecord.ListItems(cntD).SubItems(3) = rs.Fields("grNo").Value Then
                Me.listRecord.ListItems.Add , , ""
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = ""
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = ""
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) = ""
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4) = ""
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = rs.Fields("Damage").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = rs.Fields("dDate").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7) = rs.Fields("dTime").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8) = FormatNumber(rs.Fields("Cost").Value, 2)
                'MsgBox Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) & " = " & rs.Fields("TransactionID").Value
                'MsgBox "True"
            Else
                Me.listRecord.ListItems.Add , , rs.Fields("Room_no").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = rs.Fields("Company").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = rs.Fields("LastName").Value & ", " & rs.Fields("FirstName").Value & " " & rs.Fields("MiddleName").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) = rs.Fields("grNo").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4) = rs.Fields("occu_date").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = rs.Fields("Damage").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = rs.Fields("dDate").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7) = rs.Fields("dTime").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8) = FormatNumber(rs.Fields("Cost").Value, 2)
                'MsgBox Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) & " = " & rs.Fields("TransactionID").Value
                cntD = Me.listRecord.ListItems.Count
                'MsgBox "False"
            End If
        Else
                Me.listRecord.ListItems.Add , , rs.Fields("Room_no").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = rs.Fields("Company").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = rs.Fields("LastName").Value & ", " & rs.Fields("FirstName").Value & " " & rs.Fields("MiddleName").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) = rs.Fields("grNo").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4) = rs.Fields("occu_date").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = rs.Fields("Damage").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = rs.Fields("dDate").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7) = rs.Fields("dTime").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8) = FormatNumber(rs.Fields("Cost").Value, 2)
        End If
        curTotal = curTotal + Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8)
        rs.MoveNext
    Loop
    
    'If listRecord.ListItems.Count <> 0 Then
    '    For intCntX = 1 To Me.listRecord.ListItems.Count
    '        curTotal = curTotal + Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8)
    '    Next intCntX
        
            Me.listRecord.ListItems.Add , , ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8) = "-----------------"
            Me.listRecord.ListItems.Add , , ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) = ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4) = ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = ""
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7) = "Total"
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8) = FormatNumber(curTotal, 2)
    'End If
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub
Private Sub Rcounter()
    Dim intCounter As Integer
    intCounter = Me.listRecord.ListItems.Count
    Me.lblRecord.Caption = intCounter & "/" & intCounter & " Record(s)"
End Sub
