VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_Weekly_Report 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Weekly Report Manager"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9015
   ControlBox      =   0   'False
   Icon            =   "frm_Weekly_Report.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   240
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   575
      TabIndex        =   0
      Top             =   720
      Width           =   8625
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
               Picture         =   "frm_Weekly_Report.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_Weekly_Report.frx":0B24
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
            Caption         =   "&Print Report"
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
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   1111
         BorderColor     =   14215660
         Begin MSComCtl2.DTPicker dtPickerF 
            Height          =   375
            Left            =   960
            TabIndex        =   9
            Top             =   160
            Width           =   4095
            _ExtentX        =   7223
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
            Format          =   55705600
            CurrentDate     =   39746
         End
         Begin MSComCtl2.DTPicker dtPickerT 
            Height          =   375
            Left            =   5760
            TabIndex        =   11
            Top             =   160
            Width           =   4095
            _ExtentX        =   7223
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
            Format          =   55705600
            CurrentDate     =   39746
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
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
            Left            =   5280
            TabIndex        =   10
            Top             =   165
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
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
            Top             =   165
            Width           =   735
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
               Picture         =   "frm_Weekly_Report.frx":10BE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listRecord 
         Height          =   3480
         Left            =   -15
         TabIndex        =   5
         Top             =   915
         Width           =   7065
         _ExtentX        =   12462
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Day"
            Object.Width           =   2355
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2249
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
Attribute VB_Name = "frm_Weekly_Report"
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

Private Sub cmdPrint_Click()
 
    'ShellExecute Me.hwnd, "Print", App.Path & "\Reports\DRM.xls", vbNullString, App.Path & "\Reports\DRM.xls", SW_SHOWNORMAL
End Sub

Private Sub dtPicker_Change()
    LoadData
End Sub

Private Sub dtPicker_Click()
    LoadData
End Sub

Private Sub dtPickerF_Change()
    LoadData
End Sub

Private Sub dtPickerF_Click()
    LoadData
End Sub

Private Sub dtPickerT_Change()
    LoadData
End Sub

Private Sub dtPickerT_Click()
    LoadData
End Sub

Private Sub Form_Load()
    Me.dtPickerF.Value = Now
    Me.dtPickerT.Value = Now
    LoadData
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
    'Me.dtPicker.Left = Me.Width - Me.dtPicker.Width - 400


End Sub

Private Sub LoadData()
On Error GoTo err:
    Dim curTotalW As Currency
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT tblPayment.pDate, Sum(tblPayment.cPaid) AS SumOfcPaid FROM tblPayment INNER JOIN tblCustomerInfo ON tblPayment.TransactionID = tblCustomerInfo.TransactionID WHERE (((tblPayment.cPaid)>0)) GROUP BY tblPayment.pDate ORDER BY tblPayment.pDate;", cn, adOpenKeyset, adLockPessimistic
    Me.listRecord.ListItems.Clear
    Do While rs.EOF = False
        If Format(rs.Fields("pDate").Value, "MMDDYY") >= Format(Me.dtPickerF.Value, "MMDDYY") And Format(rs.Fields("pDate").Value, "MMDDYY") <= Format(Me.dtPickerT.Value, "MMDDYY") Then
            If Year(Me.dtPickerF.Value) = Year(FormatDateTime(rs.Fields("pDate").Value, vbShortDate)) Or Year(Me.dtPickerT.Value) = Year(FormatDateTime(rs.Fields("pDate").Value, vbShortDate)) Then
                Me.listRecord.ListItems.Add , , WeekdayName(Weekday(FormatDateTime(rs.Fields("pDate"), vbShortDate)))
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = rs.Fields("pDate").Value
                Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = FormatNumber(rs.Fields("SumOfcPaid").Value, 2)
                curTotalW = curTotalW + rs.Fields("SumOfcPaid").Value
            End If
        End If
    rs.MoveNext
    Loop
    If Me.listRecord.ListItems.Count <> 0 Then
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = "-------------"
        Me.listRecord.ListItems.Add , , ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(1) = ""
        Me.listRecord.ListItems(Me.listRecord.ListItems.Count).SubItems(2) = FormatNumber(curTotalW, 2)
    End If
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub
Private Sub Rcounter()
    Dim intCounter As Integer
    intCounter = Me.listRecord.ListItems.Count
    Me.lblRecord.Caption = intCounter & "/" & intCounter & " Record(s)"
End Sub
