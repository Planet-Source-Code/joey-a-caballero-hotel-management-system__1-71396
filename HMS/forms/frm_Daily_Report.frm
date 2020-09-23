VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_Daily_Report 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Daily Report Manager"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7560
   ControlBox      =   0   'False
   Icon            =   "frm_Daily_Report.frx":0000
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
               Picture         =   "frm_Daily_Report.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Daily_Report.frx":0B24
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
         Begin MSComCtl2.DTPicker dtPicker 
            Height          =   375
            Left            =   4560
            TabIndex        =   11
            Top             =   120
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   55443456
            CurrentDate     =   39746
         End
         Begin VB.ComboBox cboSearch 
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "frm_Daily_Report.frx":10BE
            Left            =   1440
            List            =   "frm_Daily_Report.frx":10C8
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00D8E9EC&
            Enabled         =   0   'False
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
            TabIndex        =   9
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Search By:"
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
            Top             =   150
            Width           =   1455
         End
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
               Picture         =   "frm_Daily_Report.frx":10E0
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
         NumItems        =   19
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
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "C/IN Time"
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "S1"
            Object.Width           =   1508
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "S2"
            Object.Width           =   1508
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Room Rate"
            Object.Width           =   2275
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Charges"
            Object.Width           =   2275
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Amount Due"
            Object.Width           =   2275
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Payment"
            Object.Width           =   2275
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Balance"
            Object.Width           =   2275
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "no Days"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Text            =   "Discount"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Text            =   "Zero-Rated"
            Object.Width           =   1984
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
Attribute VB_Name = "frm_Daily_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim dtin, dtnow As Date
Dim stime As Double
Dim intR As Integer
Dim scoTime As Double
Dim xDis As Currency
Dim xBal As Currency
Dim zRated As Currency
Dim chkSType As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Sub cboSearch_Click()
    If Me.cboSearch.Text = "" Then
        Me.txtSearch.Enabled = False
        Me.txtSearch.BackColor = &HD8E9EC
    ElseIf Me.cboSearch.Text <> "" Then
        Me.txtSearch.Enabled = True
        Me.txtSearch.BackColor = &HFFFFFF
    End If
End Sub

Private Sub cmdPrint_Click()
    expDRM
    'ShellExecute Me.hwnd, "Print", App.Path & "\Reports\DRM.xls", vbNullString, App.Path & "\Reports\DRM.xls", SW_SHOWNORMAL
End Sub




Private Sub dtPicker_Change()
    loadData
End Sub

Private Sub dtPicker_Click()
    loadData
End Sub

Private Sub Form_Load()
    Me.dtPicker.Value = Now
    loadData
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
    Me.dtPicker.Left = Me.Width - Me.dtPicker.Width - 400


End Sub

Private Sub loadData()
On Error GoTo err:
Dim sql

    'sql = " SELECT tblCustomerInfo.TransactionID, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, Sum(tblPayment.cPaid) AS SumOfcPaid, Sum(tblPayment.Meals) AS SumOfMeals, Sum(tblPayment.Damages) AS SumOfDamages, Sum(tblPayment.Services) AS SumOfServices, [SumOfMeals]+[SumOfDamages]+[SumOfServices] AS Expr1, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime FROM tblRoomRate INNER JOIN (tblCustomerInfo INNER JOIN tblPayment ON tblCustomerInfo.TransactionID = tblPayment.TransactionID) ON tblRoomRate.Room_Type = tblCustomerInfo.Room_Tariff " & _
            "GROUP BY tblCustomerInfo.TransactionID, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime;"
    
    If rs.State = adStateOpen Then rs.Close
        If chkSType = 1 Then
            On Error Resume Next
            rs.Open " SELECT tblCustomerInfo.TransactionID, tblCustomerInfo.grNo, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, Sum(tblPayment.cPaid) AS SumOfcPaid, Sum(tblPayment.Meals) AS SumOfMeals, Sum(tblPayment.Damages) AS SumOfDamages, Sum(tblPayment.Services) AS SumOfServices, Sum(tblPayment.Meals)+ Sum(tblPayment.Damages) + Sum(tblPayment.Services) AS Expr1, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime FROM tblRoomRate INNER JOIN (tblCustomerInfo INNER JOIN tblPayment ON tblCustomerInfo.TransactionID = tblPayment.TransactionID) ON tblRoomRate.Room_Type = tblCustomerInfo.Room_Tariff " & _
            "where lastname like """ & Me.txtSearch.Text & "%"" GROUP BY tblCustomerInfo.TransactionID, tblCustomerInfo.grNo, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime;", cn, adOpenKeyset, adLockPessimistic
            'rs.Open "Select * from  qryDailyReport where lastname like '" & Me.txtSearch.Text & "%';", cn, adOpenKeyset, adLockPessimistic
        Else
            On Error Resume Next
            rs.Open " SELECT tblCustomerInfo.TransactionID, tblCustomerInfo.grNo, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, Sum(tblPayment.cPaid) AS SumOfcPaid, Sum(tblPayment.Meals) AS SumOfMeals, Sum(tblPayment.Damages) AS SumOfDamages, Sum(tblPayment.Services) AS SumOfServices, Sum(tblPayment.Meals)+ Sum(tblPayment.Damages) + Sum(tblPayment.Services) AS Expr1, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime FROM tblRoomRate INNER JOIN (tblCustomerInfo INNER JOIN tblPayment ON tblCustomerInfo.TransactionID = tblPayment.TransactionID) ON tblRoomRate.Room_Type = tblCustomerInfo.Room_Tariff " & _
            "where Company like """ & Me.txtSearch.Text & "%"" GROUP BY tblCustomerInfo.TransactionID, tblCustomerInfo.grNo, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime;", cn, adOpenKeyset, adLockPessimistic
            
            'rs.Open "Select * from qryDailyReport where Company like '" & Me.txtSearch.Text & "%';", cn, adOpenKeyset, adLockPessimistic
        End If
       
    Me.listRecord.ListItems.Clear
    Do While rs.EOF = False
     'MsgBox Format(rs.Fields("occu_date").Value, "MMDDYY") & " <= " & Format(Me.dtPicker.Value, "MMDDYY")
        If rs.Fields("Stat").Value <> "R" Then
                If Format(rs.Fields("occu_date").Value, "MMDDYY") <= Format(Me.dtPicker.Value, "MMDDYY") Then
                    Me.listRecord.ListItems.Add , , rs.Fields("Room_no").Value
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = rs.Fields("LastName").Value & ", " & rs.Fields("FirstName").Value & " " & rs.Fields("MiddleName").Value
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = rs.Fields("Company").Value
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) = rs.Fields("grNo").Value
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8) = FormatNumber(rs.Fields("Room_Rate").Value, 2)
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(11) = FormatNumber(rs.Fields("SumOfcPaid").Value, 2)
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4) = rs.Fields("occu_date").Value
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = rs.Fields("occu_time").Value
                    
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(9) = FormatNumber(rs.Fields("Expr1").Value, 2)
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(15) = rs.Fields("TransactionID").Value
                    
                    If rs.Fields("coTime").Value = "0:00" Then
                        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(16) = FormatDateTime(Now, vbShortTime)
                        scoTime = Format(FormatDateTime(Now, vbShortTime), "HHMMSS")
                    Else
                        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(16) = rs.Fields("coTime").Value
                        scoTime = Format(rs.Fields("coTime").Value, "HHMMSS")
                    End If
                    
                    If rs.Fields("coDate") = "0" Then
                        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(13) = FormatDateTime(Now, vbShortDate)
                    Else
                        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(13) = rs.Fields("coDate").Value
                    End If
                    
                    stime = Format(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5), "HHMMSS")
                    dtin = Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4)
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7) = rs.Fields("Stat").Value
                    If UCase(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7)) = "C/OUT" Then
                        getCONDay
                    Else
                        getDays
                    End If
                    
                    getCODay
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = "R" & intR
                    Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(10) = FormatNumber(rs.Fields("Room_Rate").Value * Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) + rs.Fields("Expr1").Value, 2)
                    getDiscount
                    xDis = (CCur(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8)) * CDbl(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14)) / 100) * Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(17) + rs.Fields("Expr1").Value
                    'MsgBox xDis
                    
                    If Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(18) = "Yes" Then
                        zRated = FormatNumber((CCur(rs.Fields("Room_Rate").Value) * CDbl(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14)) / 1.12) + rs.Fields("Expr1").Value, 2)
                        'MsgBox zRated
                        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(12) = FormatNumber(CCur(zRated) - xDis - rs.Fields("SumOfcPaid").Value + rs.Fields("Expr1").Value, 2)
                    Else
                        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(12) = FormatNumber(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(10) - xDis - rs.Fields("SumOfcPaid").Value + rs.Fields("Expr1").Value, 2)
                    End If
             End If
                'Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(12) = FormatNumber(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(10) - xDis - rs.Fields("SumOfcPaid").Value, 2)
        End If
        rs.MoveNext
    Loop
    Exit Sub

err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub getDays()
    
    dtnow = FormatDateTime(Now, vbShortDate)
    ' Check-in form 12:00am - 5:00am
    ' Code Ok
    ' For Adjustmet
    If stime >= 100 And stime <= 50000 Then
        If DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1.5
        ElseIf DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 2
        Else
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
        End If
    ' Check-in form 5:01am - 7:00am
    ' Code Ok
    ' For Adjustmet
    ElseIf stime >= 50100 And stime <= 70000 Then
        If DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") <= 120100 Then 'and Format(Time, "HHMMSS") <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
        
        ElseIf DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") >= 120100 Then 'and Format(Time, "HHMMSS") <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
                   
        ElseIf DateDiff("d", dtin, dtnow) > 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
             
        ElseIf DateDiff("d", dtin, dtnow) > 0 And Format(Time, "HHMMSS") >= 140100 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
            
        Else
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow)
        End If
    ' Will Follow
    ' Code Ok
    ' For Adjustmet
    ElseIf DateDiff("d", dtin, dtnow) = 0 And stime >= 70000 And stime <= 120000 Then
        If DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
        ElseIf DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") >= 140100 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
        End If
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") >= 140100 Then
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") <= 120100 Then
    
       Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow)
    Else
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = 1
    End If
        'gDays = Me.cboDays.Text
End Sub

Private Sub listRecord_Click()
On Error GoTo err:
    Dim intCounter, intSelected As Integer
    intCounter = Me.listRecord.ListItems.Count
    intSelected = Me.listRecord.SelectedItem.Index
    Me.lblRecord.Caption = intSelected & "/" & intCounter & " Record(s)"
    Exit Sub
err:
    Exit Sub
End Sub

Private Sub listRecord_DblClick()
    transIDs = Me.listRecord.SelectedItem.SubItems(15)
    Dim sql
    Dim rsTrans As ADODB.Recordset
    Set rsTrans = New ADODB.Recordset
    
    With rsTrans
        If .State = adStateOpen Then .Close
        
        
        .Open "SELECT tblCustomerInfo.TransactionID, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, Sum(tblPayment.cPaid) AS SumOfcPaid, Sum(tblPayment.Meals) AS SumOfMeals, Sum(tblPayment.Damages) AS SumOfDamages, Sum(tblPayment.Services) AS SumOfServices, Sum(tblPayment.Meals)+ Sum(tblPayment.Damages) + Sum(tblPayment.Services) AS Expr1, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime FROM tblRoomRate INNER JOIN (tblCustomerInfo INNER JOIN tblPayment ON tblCustomerInfo.TransactionID = tblPayment.TransactionID) ON tblRoomRate.Room_Type = tblCustomerInfo.Room_Tariff WHERE (((tblCustomerInfo.TransactionID)=" & transIDs & ")) GROUP BY tblCustomerInfo.TransactionID, " & _
        "tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, " & _
        "tblCustomerInfo.occu_time , tblRoomRate.Room_Rate, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime;", cn, adOpenKeyset, adLockPessimistic

        '.Open "SELECT tblCustomerInfo.TransactionID, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, Sum(tblPayment.cPaid) AS SumOfcPaid, Sum(tblPayment.Meals) AS SumOfMeals, Sum(tblPayment.Damages) AS SumOfDamages, Sum(tblPayment.Services) AS SumOfServices, Sum(tblPayment.Meals)+ Sum(tblPayment.Damages) + Sum(tblPayment.Services) AS Expr1, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime " & _
              "FROM tblRoomRate INNER JOIN (tblCustomerInfo INNER JOIN tblPayment ON tblCustomerInfo.TransactionID = tblPayment.TransactionID) ON tblRoomRate.Room_Type = tblCustomerInfo.Room_Tariff WHERE tblCustomerInfo.TransactionID = " & transIDs & ";", cn, adOpenKeyset, adLockPessimistic
              '" GROUP BY tblCustomerInfo.TransactionID, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.Company, tblCustomerInfo.Room_no, tblCustomerInfo.occu_date, tblCustomerInfo.occu_time, tblRoomRate.Room_Rate, tblCustomerInfo.Stat, tblCustomerInfo.Room_Tariff, tblCustomerInfo.coDate, tblCustomerInfo.coTime;", cn, adOpenKeyset, adLockPessimistic

        '.Open "SELECT * FROM qryDailyReport WHERE TransactionID = " & transIDs, cn, adOpenKeyset, adLockOptimistic
        '.Open sql & " WHERE TransactionID = " & transIDs, cn, adOpenKeyset, adLockPessimistic
           sName = .Fields("LastName") & ", " & .Fields("FirstName") & " " & .Fields("MiddleName")
           iRoomNum = .Fields("Room_no")
           sRoomTar = .Fields("Room_tariff")
           curRoomRate = .Fields("Room_Rate")
    End With
    Set rsTrans = Nothing
    
    EXReport
End Sub

Private Sub EXReport()
    
On Error GoTo err:
    Dim conn As New ADODB.Connection
    Dim xlApp As New Excel.Application
    Dim xlwk As New Excel.Workbook
    Dim ctr As Integer
    
    xlApp.Interactive = True
    Set xlwk = xlApp.Workbooks.Open(App.Path & "\Reports\TransSum.xls")
    
    'ctr = 5 ' start data after headings
    'xlApp.Workbooks.Add
     
    xlApp.Cells(4, 2) = sName
    xlApp.Cells(4, 8) = Me.listRecord.SelectedItem.SubItems(1) 'Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1)
    xlApp.Cells(4, 12) = Me.listRecord.SelectedItem.SubItems(10) ' Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(10)
    
    xlApp.Cells(5, 2) = iRoomNum
    xlApp.Cells(5, 7) = sRoomTar
    xlApp.Cells(5, 11) = curRoomRate
    
    xlApp.Cells(7, 3) = Me.listRecord.SelectedItem.SubItems(4) 'Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4)
    xlApp.Cells(7, 11) = Me.listRecord.SelectedItem.SubItems(14) 'Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14)
        
   
    Dim rsTrans As ADODB.Recordset
    Set rsTrans = New ADODB.Recordset
    ctr = 11
    With rsTrans
    
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM tblAddMeals WHERE TransID = " & transIDs, cn, adOpenKeyset, adLockOptimistic
        Do While rsTrans.EOF = False
            ctr = ctr + 1
            xlApp.Cells(ctr, 1) = rsTrans("oDate").Value
            xlApp.Cells(ctr, 2) = rsTrans("MType").Value
            xlApp.Cells(ctr, 4) = rsTrans("MCost").Value
            rsTrans.MoveNext
        Loop
    End With
    Set rsTrans = Nothing
    
    Set rsTrans = New ADODB.Recordset
    ctr = 11
    With rsTrans
    
    If .State = adStateOpen Then .Close
        .Open "SELECT * FROM tblAddOnDamage WHERE TransID = " & transIDs, cn, adOpenKeyset, adLockOptimistic
        Do While rsTrans.EOF = False
            ctr = ctr + 1
            xlApp.Cells(ctr, 10) = rsTrans("dDate").Value
            xlApp.Cells(ctr, 11) = rsTrans("Damage").Value
            xlApp.Cells(ctr, 12) = rsTrans("Cost").Value
            rsTrans.MoveNext
        Loop
    End With
    Set rsTrans = Nothing
    
    Set rsTrans = New ADODB.Recordset
    ctr = 11
    With rsTrans
    
    If .State = adStateOpen Then .Close
        .Open "SELECT * FROM tblAddOthers WHERE TransID = " & transIDs, cn, adOpenKeyset, adLockOptimistic
        Do While rsTrans.EOF = False
            ctr = ctr + 1
            xlApp.Cells(ctr, 6) = rsTrans("oDate").Value
            xlApp.Cells(ctr, 7) = rsTrans("Services").Value
            xlApp.Cells(ctr, 8) = rsTrans("Price").Value
            rsTrans.MoveNext
        Loop
    End With
    Set rsTrans = Nothing
    
    
    
    'xlApp.Range("A" & Trim(Str(ctr))).Value = rs("customerID")
    
    'Do While Not rs.EOF
    '    ctr = ctr + 1
    '    xlApp.Range("A" & Trim(Str(ctr))).Value = rs("customerID")
    '    xlApp.Range("B" & Trim(Str(ctr))).Value = rs("companyName")
    '    xlApp.Range("C" & Trim(Str(ctr))).Value = rs("contactName")
    '    xlApp.Range("D" & Trim(Str(ctr))).Value = rs("address")
    '    xlApp.Range("E" & Trim(Str(ctr))).Value = rs("city")
    '    xlApp.Range("F" & Trim(Str(ctr))).Value = rs("phone")
    '    rs.MoveNext
    'Loop
    
    xlApp.Visible = True
    xlApp.ActiveWindow.SelectedSheets.PrintPreview
    xlApp.ActiveWorkbook.Saved = True
    'xlApp.ActiveWindow
    'xlApp.ActiveWindow.Close (False)
    xlApp.Quit
    Exit Sub
err:
    MsgBox err.Description, vbCritical

End Sub


Private Sub getCODay()
If UCase(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7)) = "C/OUT" Then
    dtnow = Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(13)

Else
    dtnow = FormatDateTime(Now, vbShortDate)

End If
    
    If stime >= 100 And stime <= 50000 Then
        If DateDiff("d", dtin, dtnow) = 0 And scoTime <= 120100 Then
            'MsgBox "1"
            intR = DateDiff("d", dtin, dtnow) + 1
        ElseIf DateDiff("d", dtin, dtnow) = 0 And scoTime >= 120100 Then
            'MsgBox "2"
            intR = DateDiff("d", dtin, dtnow) + 2
        ElseIf DateDiff("d", dtin, dtnow) <> 0 And scoTime <= 120100 Then
            intR = DateDiff("d", dtin, dtnow) + 1
            'MsgBox "3"
        Else 'If DateDiff("d", dtin, dtnow) <> 0 And scoTime >= 120100 Then
            intR = DateDiff("d", dtin, dtnow) + 2
            'MsgBox "4"
        End If
    Else
        If DateDiff("d", dtin, dtnow) = 0 Then 'And Format(Time, "HHMMSS") <= 120100 Then
            intR = DateDiff("d", dtin, dtnow) + 1
            'MsgBox "5"
        ElseIf DateDiff("d", dtin, dtnow) <> 0 And scoTime <= 120000 Then
            intR = DateDiff("d", dtin, dtnow)
            'MsgBox "6"
            If UCase(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7)) = "C/OUT" Then
                intR = DateDiff("d", dtin, dtnow)
            'MsgBox "7"
            End If
        ElseIf DateDiff("d", dtin, dtnow) <> 0 And scoTime >= 120100 Then
            If UCase(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7)) = "C/OUT" Then
                intR = DateDiff("d", dtin, dtnow) + 1
                'MsgBox "8"
            Else
                intR = DateDiff("d", dtin, dtnow) + 1
                'MsgBox Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(7)
           
                'MsgBox "9"
            End If
        'else if DateDiff("d", dtin, dtnow) <> 0 And scoTime <= 120100 Then
        End If
    End If
End Sub

Private Sub getCONDay()
    dtnow = Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(13)
    
    If stime >= 100 And stime <= 50000 Then
        If DateDiff("d", dtin, dtnow) >= 0 And scoTime >= 120100 And scoTime <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1.5
        ElseIf DateDiff("d", dtin, dtnow) >= 0 And scoTime >= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 2
        Else
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
        End If
    ' Check-in form 5:01am - 7:00am
    ' Code Ok
    ' For Adjustmet
    ElseIf stime >= 50100 And stime <= 70000 Then
        If DateDiff("d", dtin, dtnow) = 0 And scoTime <= 120100 Then 'and Format(Time, "HHMMSS") <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
        
        ElseIf DateDiff("d", dtin, dtnow) = 0 And scoTime >= 120100 Then 'and Format(Time, "HHMMSS") <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
                   
        ElseIf DateDiff("d", dtin, dtnow) > 0 And scoTime >= 120100 And scoTime <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
             
        ElseIf DateDiff("d", dtin, dtnow) > 0 And scoTime >= 140100 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
            
        Else
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow)
        End If
    ' Will Follow
    ' Code Ok
    ' For Adjustmet
    ElseIf DateDiff("d", dtin, dtnow) = 0 And stime >= 70000 And stime <= 120000 Then
        If DateDiff("d", dtin, dtnow) = 0 And scoTime <= 140000 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
        ElseIf DateDiff("d", dtin, dtnow) = 0 And scoTime >= 140100 Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
        End If
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And scoTime >= 120100 And scoTime <= 140000 Then
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 0.5
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And scoTime >= 140100 Then
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow) + 1
    ElseIf DateDiff("d", dtin, dtnow) <> 0 And scoTime <= 120100 Then
    
       Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = DateDiff("d", dtin, dtnow)
    Else
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14) = 1
    End If
End Sub

Private Sub Rcounter()
    Dim intCounter As Integer
    intCounter = Me.listRecord.ListItems.Count
    Me.lblRecord.Caption = intCounter & "/" & intCounter & " Record(s)"
End Sub

Private Sub expDRM()
    Dim xlApp As New Excel.Application
    Dim xlwk As New Excel.Workbook
    Dim ctr As Integer
    Dim i As Integer
    xlApp.Interactive = True
    Set xlwk = xlApp.Workbooks.Open(App.Path & "\Reports\DRM.xls")
    ctr = 4
    For i = 1 To Me.listRecord.ListItems.Count
        ctr = ctr + 1
        xlApp.Cells(ctr, 1) = Me.listRecord.ListItems(i).Text
        xlApp.Cells(ctr, 2) = Me.listRecord.ListItems(i).SubItems(1)
        xlApp.Cells(ctr, 3) = Me.listRecord.ListItems(i).SubItems(2)
        xlApp.Cells(ctr, 4) = Me.listRecord.ListItems(i).SubItems(3)
        xlApp.Cells(ctr, 5) = Me.listRecord.ListItems(i).SubItems(4)
        xlApp.Cells(ctr, 6) = Me.listRecord.ListItems(i).SubItems(5)
        xlApp.Cells(ctr, 7) = Me.listRecord.ListItems(i).SubItems(6)
        xlApp.Cells(ctr, 8) = Me.listRecord.ListItems(i).SubItems(7)
        xlApp.Cells(ctr, 9) = Me.listRecord.ListItems(i).SubItems(8)
        xlApp.Cells(ctr, 10) = Me.listRecord.ListItems(i).SubItems(9)
        xlApp.Cells(ctr, 11) = Me.listRecord.ListItems(i).SubItems(10)
        xlApp.Cells(ctr, 12) = Me.listRecord.ListItems(i).SubItems(11)
        xlApp.Cells(ctr, 13) = Me.listRecord.ListItems(i).SubItems(12)
    Next i
   
        xlApp.Visible = True
        xlApp.ActiveWindow.SelectedSheets.PrintPreview
        xlApp.ActiveWorkbook.Saved = True
        xlApp.Quit
End Sub

Private Sub txtSearch_Change()
    If Me.cboSearch.Text = "Last Name" Then
        chkSType = 1
    Else
        chkSType = 2
    End If
    
    'If FormatDateTime(Me.dtPicker.Value, vbShortDate) = FormatDateTime(Now, vbShortDate) Then
        loadData
    'Else
    '    dtPicker_Change
    'End If
    If Me.listRecord.ListItems.Count <> 0 Then Me.listRecord.Enabled = True
    If Me.listRecord.ListItems.Count = 0 Then Me.listRecord.Enabled = False
End Sub


Private Sub getDiscount()
On Error GoTo err:
        If rs1.State = adStateOpen Then rs1.Close
        rs1.Open "Select * from tblPayment where TransactionID = " & Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(15) & ";", cn, adOpenKeyset, adLockPessimistic
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(17) = rs1.Fields("Discount").Value
        If rs1.Fields("isNonTax").Value = "true" Then
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(18) = "Yes"
        Else
            Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(18) = "No"
        End If
        
    Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

'Private Sub calDiscount()
'    xDis = (CCur(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(8)) * CDbl(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(14)) / 100) * CCur(Me.txtDiscount.Text)
'    'FormatNumber(Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(10) - rs.Fields("SumOfcPaid").Value, 2)
'End Sub
