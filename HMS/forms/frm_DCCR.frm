VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DCCR 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Daily Cash Collection Report"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7560
   ControlBox      =   0   'False
   Icon            =   "frm_DCCR.frx":0000
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
               Picture         =   "frm_DCCR.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_DCCR.frx":0B24
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
               Picture         =   "frm_DCCR.frx":10BE
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Company  Name"
            Object.Width           =   3307
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer's Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "GR #"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Date Paid"
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Time Paid"
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Payment Type"
            Object.Width           =   2275
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Payment"
            Object.Width           =   2275
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
Attribute VB_Name = "frm_DCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
Dim rs As New ADODB.Recordset
Dim dtin, dtnow As Date
Dim stime As Double
Dim intR As Integer
Dim scoTime As Double

Dim chkSType As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Sub cmdPrint_Click()
    'expDRM
    'ShellExecute Me.hwnd, "Print", App.Path & "\Reports\DRM.xls", vbNullString, App.Path & "\Reports\DRM.xls", SW_SHOWNORMAL
    Dim sql
        If rs.State = adStateOpen Then rs.Close

        sql = "SELECT tblCustomerInfo.Company, tblCustomerInfo!LastName+ ', '+tblCustomerInfo!FirstName+' '+tblCustomerInfo!MiddleName AS CName, tblCollection.dDate, tblCollection.dTime, tblCollection.Amount, tblCustomerInfo.TransactionID FROM tblCustomerInfo INNER JOIN tblCollection ON tblCustomerInfo.TransactionID = tblCollection.TransID GROUP BY tblCustomerInfo.Company, tblCustomerInfo!LastName+', '+tblCustomerInfo!FirstName+' '+tblCustomerInfo!MiddleName, tblCollection.dDate, tblCollection.dTime, tblCollection.Amount, tblCustomerInfo.TransactionID;"

        rs.Open sql, cn, adOpenKeyset, adLockPessimistic
        
        Set DCCR.DataSource = rs
        DCCR.PrintReport , rptRangeAllPages
End Sub

Private Sub Form_Load()
DataLoader
Rcounter
End Sub

Private Sub Form_Resize()
    ReArrangeControls
End Sub

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

Private Sub listRecord_Click()
On Error GoTo err:
    Dim intCounter, intSelected As Integer
    intCounter = Me.listRecord.ListItems.Count - 3
    intSelected = Me.listRecord.SelectedItem.Index
    If intSelected > intCounter Then intSelected = intCounter
    Me.lblRecord.Caption = intSelected & "/" & intCounter & " Record(s)"
    Exit Sub
err:
    Exit Sub
End Sub

Private Sub Rcounter()
    Dim intCounter As Integer
    intCounter = Me.listRecord.ListItems.Count - 3
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


Private Sub DataLoader()
On Error GoTo err:
    Dim sql
    Dim intCnt As Integer
    Dim curTotal As Double
    Dim CurCashInHand As Currency
    curTotal = 0
    If rs.State = adStateOpen Then rs.Close

        sql = "SELECT tblCustomerInfo.Company, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.TransactionID, tblCollection.dDate, tblCollection.dTime, tblCollection.PType, tblCollection.Amount FROM tblCustomerInfo INNER JOIN tblCollection ON tblCustomerInfo.TransactionID = tblCollection.TransID GROUP BY tblCustomerInfo.Company, tblCustomerInfo.LastName, tblCustomerInfo.FirstName, tblCustomerInfo.MiddleName, tblCustomerInfo.TransactionID, tblCollection.dDate, tblCollection.dTime, tblCollection.PType, tblCollection.Amount;"
        rs.Open sql, cn, adOpenKeyset, adLockPessimistic
        
    Do While rs.EOF = False
        Me.listRecord.ListItems.Add , , rs.Fields("Company").Value
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = rs.Fields("LastName").Value & ", " & rs.Fields("FirstName").Value & " " & rs.Fields("MiddleName").Value
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = rs.Fields("TransactionID").Value
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(3) = rs.Fields("dDate").Value
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(4) = rs.Fields("dTime").Value
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = rs.Fields("PType").Value
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = FormatNumber(rs.Fields("Amount").Value, 2)

        curTotal = curTotal + FormatNumber(rs.Fields("Amount").Value, 2)

        If rs.Fields("PType").Value = "Cash" Then
            CurCashInHand = CurCashInHand + FormatNumber(rs.Fields("Amount").Value, 2)
        End If
        rs.MoveNext
    Loop
    
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = "------------------"
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = "Total"
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = FormatNumber(curTotal, 2)
        
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = "------------------"
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(5) = "Cash-In-Hand"
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(6) = FormatNumber(CurCashInHand, 2)
        
Exit Sub
err:
    MsgBox err.Description, vbCritical
    
End Sub
