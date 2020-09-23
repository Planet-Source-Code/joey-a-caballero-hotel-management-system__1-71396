VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Welcome 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Welcome to BPH System"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7560
   ControlBox      =   0   'False
   Icon            =   "frm_Welcome.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Welcome.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Welcome.frx":2A1AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Welcome.frx":54E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Welcome.frx":7EAA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Welcome.frx":7F382
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox bgMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   0
      Top             =   480
      Width           =   7065
      Begin MSComctlLib.ListView lvr 
         Height          =   3495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6165
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   128
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frm_Welcome.frx":A8FA4
         OLEDragMode     =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "1"
            Object.Width           =   6615
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "2"
            Object.Width           =   6615
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "3"
            Object.Width           =   6615
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "4"
            Object.Width           =   6615
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "5"
            Object.Width           =   6615
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "6"
            Object.Width           =   6615
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   6615
         EndProperty
      End
   End
   Begin HMS.b8Container b8cMain 
      Height          =   5940
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10478
      BorderColor     =   12632256
      InsideBorderColor=   14215660
      ShadowColor1    =   14215660
      ShadowColor2    =   14215660
   End
End
Attribute VB_Name = "frm_Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rsReserved As New ADODB.Recordset
Dim rsUpdate As New ADODB.Recordset

Private Sub Form_Resize()
    ReArrangeControls
End Sub
Public Function Form_Find()
    frmFindListItem.ShowFind lvr
End Function

Private Sub Form_Activate()
    mdiMain.RegMDIChild Me
    Me.WindowState = vbMaximized
    loadRooms
End Sub

Private Sub ReArrangeControls()
On Error Resume Next
    Me.ScaleMode = vbPixels
    b8cMain.Move Form_LeftMargin - 3, Form_TopMargin - 3, Me.ScaleWidth - (Form_LeftMargin - 3) * 2, Me.ScaleHeight - (Form_TopMargin - 3) * 2
    
    bgMain.Move Form_LeftMargin, Form_TopMargin, Me.ScaleWidth - Form_LeftMargin * 2, Me.ScaleHeight - Form_TopMargin * 2
    
    lvr.Move lvr.Left, 0, bgMain.Width - (lvr.Left * 2)
    lvr.Height = bgMain.Height
End Sub

Public Sub loadRooms()
On Error GoTo err:
    Dim rDate As Date
    rDate = Format(Date, "MMDDYY")
    Dim rRoNo As Integer
    
    Me.lvr.ListItems.Clear
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblroom", cn, adOpenKeyset, adLockPessimistic
        
    Do While rs.EOF = False
        
        If rs.Fields("occupied").Value = "false" Then
            rRoNo = rs.Fields("Room_no").Value
            
            If rsReserved.State = adStateOpen Then rsReserved.Close
            rsReserved.Open "Select * from tblReserved where rNo=" & rRoNo & ";", cn, adOpenKeyset, adLockPessimistic
            'MsgBox rRoNo
            Do While rsReserved.EOF = False
                If Format(Date, "MMDDYY") >= Format(rsReserved.Fields("rDate").Value, "MMDDYY") And Format(Date, "MMDDYY") <= Format(rsReserved.Fields("uDate").Value, "MMDDYY") Then
                    'If rsUpdate.State = adStateOpen Then rsUpdate.Close
                    '    rsUpdate.Open "Select * from tblroom where Room_no = " & rRoNo & ";", cn, adOpenKeyset, adLockPessimistic
                    '    rsUpdate.Fields("occupied").Value = "reserved"
                    '    rsUpdate.Update
                          Me.lvr.ListItems.Add , , "Room " & rs.Fields("Room_no").Value & " " & rs.Fields("Room_tariff").Value, 3
                    Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(5) = rsReserved.Fields("transID").Value
                    Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(1) = 2
                    Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(2) = rs.Fields("Floor_no").Value
                    Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(3) = rs.Fields("Room_tariff").Value
                    Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(4) = rs.Fields("Room_no").Value
                    Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(6) = "True"
                End If
            rsReserved.MoveNext
            Loop
          '  MsgBox "False"
            On Error GoTo nextstep:
            If Me.lvr.ListItems(Me.lvr.ListItems.Count).SubItems(4) = rRoNo Then
            Else
nextstep:
            Me.lvr.ListItems.Add , , "Room " & rs.Fields("Room_no").Value & " " & rs.Fields("Room_tariff").Value, 1
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(1) = 0
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(2) = rs.Fields("Floor_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(3) = rs.Fields("Room_tariff").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(4) = rs.Fields("Room_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(5) = rs.Fields("transID").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(6) = "False"
            End If
            
        ElseIf rs.Fields("occupied").Value = "true" Then
            Me.lvr.ListItems.Add , , "Room " & rs.Fields("Room_no").Value & " " & rs.Fields("Room_tariff").Value, 2
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(1) = 1
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(2) = rs.Fields("Floor_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(3) = rs.Fields("Room_tariff").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(4) = rs.Fields("Room_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(5) = rs.Fields("transID").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(6) = "False"
       ' ElseIf rs.Fields("occupied").Value = "reserved" Then
       '     If rsReserved.State = adStateOpen Then rsReserved.Close
       '     rsReserved.Open "Select * from tblReserved where rNo=" & rRoNo & ";", cn, adOpenKeyset, adLockPessimistic
       ElseIf rs.Fields("occupied").Value = "Storage" Then
            Me.lvr.ListItems.Add , , "Room " & rs.Fields("Room_no").Value & " " & "Storage Room", 4
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(1) = 3
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(2) = rs.Fields("Floor_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(3) = rs.Fields("Room_tariff").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(4) = rs.Fields("Room_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(5) = rs.Fields("transID").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(6) = "False"
       ElseIf rs.Fields("occupied").Value = "Dirty" Then
            Me.lvr.ListItems.Add , , "Room " & rs.Fields("Room_no").Value & " " & rs.Fields("Room_tariff").Value, 5
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(1) = 4
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(2) = rs.Fields("Floor_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(3) = rs.Fields("Room_tariff").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(4) = rs.Fields("Room_no").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(5) = rs.Fields("transID").Value
            Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(6) = "False"
        End If
    rs.MoveNext
    Loop
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub lvr_Click()
    If Me.lvr.SelectedItem.SubItems(6) = "True" Then
        isRchkin = True
    Else
        isRchkin = False
    End If
End Sub

Private Sub LVR_DblClick()
    chkin
End Sub

Private Sub lvr_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        transID = Me.lvr.SelectedItem.SubItems(5)
'    If Me.lvr.SelectedItem.SubItems(6) = "True" Then
'        isRchkin = True
'        'MsgBox isRchkin
'    Else
'        isRchkin = False
'        'MsgBox isRchkin
'    End If
    
    Dim rDate As Date
    rDate = Format(Date, "MMDDYY")
    isRoomNo = Me.lvr.SelectedItem.SubItems(4)
    If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblroom where transID = " & Me.lvr.SelectedItem.SubItems(5) & ";", cn, adOpenKeyset, adLockPessimistic
    
    If Me.lvr.SelectedItem.SubItems(1) = 0 Then '''''vacant
        mdiMain.mnupcheckout.Enabled = False
        mdiMain.mnupFChkOut.Enabled = False
        mdiMain.mnupcheckin.Enabled = True
        isPayment = False
        mdiMain.mnupreservation.Enabled = True
        mdiMain.mnupcancel.Enabled = True
        mdiMain.mnupreserve.Enabled = True
        mdiMain.mnucPayment.Enabled = False
        mdiMain.mnuDamages.Enabled = False
        mdiMain.mnuAddServ.Enabled = False
        mdiMain.mnuMeals.Enabled = False
    ElseIf Me.lvr.SelectedItem.SubItems(1) = 1 Then   '''''occupied
        mdiMain.mnupcheckin.Enabled = False
        mdiMain.mnupcheckout.Enabled = True
        mdiMain.mnupFChkOut.Enabled = True
        mdiMain.mnupreservation.Enabled = True
        mdiMain.mnucPayment.Enabled = True
        mdiMain.mnupcancel.Enabled = True
        mdiMain.mnuDamages.Enabled = True
        mdiMain.mnuAddServ.Enabled = True
        mdiMain.mnuMeals.Enabled = True
    ElseIf Me.lvr.SelectedItem.SubItems(1) = 2 Then     '''''reserved
        mdiMain.mnupcheckout.Enabled = False
        mdiMain.mnupFChkOut.Enabled = False
        mdiMain.mnupreservation.Enabled = True
        mdiMain.mnupreserve.Enabled = True
        mdiMain.mnupcancel.Enabled = True
        mdiMain.mnucPayment.Enabled = False
        mdiMain.mnuDamages.Enabled = False
        mdiMain.mnuAddServ.Enabled = False
        mdiMain.mnuMeals.Enabled = False
      '  If rDate = rs.Fields("dReserved").Value Then
            mdiMain.mnupcheckin.Enabled = True
      '  Else
       '     mdiMain.mnupcheckin.Enabled = True
      '  End If
    ElseIf Me.lvr.SelectedItem.SubItems(1) = 3 Then
        Exit Sub
    ElseIf Me.lvr.SelectedItem.SubItems(1) = 4 Then
        GoTo pop:
    End If
    If Button = vbRightButton Then
        PopupMenu mdiMain.mnupopup
    End If
    Exit Sub
pop:
    If Button = vbRightButton Then
        PopupMenu mdiMain.mnupopAvailable
    End If
End Sub
Public Sub chkin()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblCustomerInfo where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
    
   If Me.lvr.SelectedItem.SubItems(1) = 0 Then
    roomtype = Me.lvr.SelectedItem.SubItems(3)
    isPayment = False
        With frm_checkin
            isRInfo = False
            .lblFloor.Caption = Me.lvr.SelectedItem.SubItems(2)
            .lblroomno.Caption = Me.lvr.SelectedItem.SubItems(4)
            .lblRT.Caption = Me.lvr.SelectedItem.SubItems(3)
            'isPayment = False
            .Show
        End With
            mdiMain.Enabled = False
    ElseIf Me.lvr.SelectedItem.SubItems(1) = 1 Then
        'isPayment = False
        MsgBox "This Room is Occupied by Mr./Mrs. " & rs.Fields("FirstName").Value & " " & rs.Fields("LastName").Value & vbCrLf & vbCrLf & "Thank You", vbInformation, "Room Info"
    ElseIf Me.lvr.SelectedItem.SubItems(1) = 2 Then
        roomtype = Me.lvr.SelectedItem.SubItems(3)
        isPayment = False
        isRInfo = True
            With frm_checkin
                .lblFloor.Caption = Me.lvr.SelectedItem.SubItems(2)
                .lblroomno.Caption = Me.lvr.SelectedItem.SubItems(4)
                .lblRT.Caption = Me.lvr.SelectedItem.SubItems(3)
                'isPayment = False
                .lblroomno.Caption = rs.Fields("Room_no").Value
                .txtcompany.Text = rs.Fields("Company").Value
                .txtlname.Text = rs.Fields("LastName").Value
                .txtfname.Text = rs.Fields("FirstName").Value
                .txtmname.Text = rs.Fields("MiddleName").Value
       
                If rs.Fields("ContactNo").Value = "N/A" Or rs.Fields("ContactNo").Value = "" Then
                     .txtcontactnoCode.Text = "N/A"
                     .txtcontactno.Text = "N/A"
                Else
                     .txtcontactnoCode.Text = Split(rs.Fields("ContactNo").Value, "-")(0)
                     .txtcontactno.Text = Split(rs.Fields("ContactNo").Value, "-")(1)
                End If
                
                .txtemail.Text = rs.Fields("Email").Value
                .cbo_nationality.Text = rs.Fields("Nationality").Value
                .txthomeadd.Text = rs.Fields("HomeAdd").Value
                dUnlock
                .Show
                
            End With
                mdiMain.Enabled = False
        
        
        'isPayment = False
         '   frm_Payment.Show
         'MsgBox "Ready for coding"
    End If
    
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Public Sub chkReserved()
   'If Me.lvr.SelectedItem.SubItems(1) = 0 Then
        roomtype = Me.lvr.SelectedItem.SubItems(3)
        With frm_Reserved
            .lblFloor.Caption = Me.lvr.SelectedItem.SubItems(2)
            .lblroomno.Caption = Me.lvr.SelectedItem.SubItems(4)
            .lblRT.Caption = Me.lvr.SelectedItem.SubItems(3)
            'isPayment = False
            .Show
        End With
    'End If
End Sub


Public Sub VloadRooms()
On Error GoTo err:
    Dim rDate As Date
    rDate = Format(Date, "MMDDYY")
    Dim rRoNo As Integer
    
    Me.lvr.ListItems.Clear
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblroom where occupied = 'false'", cn, adOpenKeyset, adLockPessimistic
        
    Do While rs.EOF = False
        Me.lvr.ListItems.Add , , "Room " & rs.Fields("Room_no").Value & " " & rs.Fields("Room_tariff").Value, 1
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(1) = 0
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(2) = rs.Fields("Floor_no").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(3) = rs.Fields("Room_tariff").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(4) = rs.Fields("Room_no").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(5) = rs.Fields("transID").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(6) = "False"
        rs.MoveNext
    Loop
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub OloadRooms()
On Error GoTo err:
    Dim rDate As Date
    rDate = Format(Date, "MMDDYY")
    Dim rRoNo As Integer
    
    Me.lvr.ListItems.Clear
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblroom where occupied = 'true'", cn, adOpenKeyset, adLockPessimistic
        
    Do While rs.EOF = False
        Me.lvr.ListItems.Add , , "Room " & rs.Fields("Room_no").Value & " " & rs.Fields("Room_tariff").Value, 2
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(1) = 1
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(2) = rs.Fields("Floor_no").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(3) = rs.Fields("Room_tariff").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(4) = rs.Fields("Room_no").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(5) = rs.Fields("transID").Value
        Me.lvr.ListItems.Item(Me.lvr.ListItems.Count).SubItems(6) = "False"
        rs.MoveNext
    Loop
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub dUnlock()
    With frm_checkin
        .txtlname.Locked = False
        .txtlname.BackColor = vbWhite
        .txtfname.Locked = False
        .txtfname.BackColor = vbWhite
        .txtmname.Locked = False
        .txtmname.BackColor = vbWhite
        .txtcompany.Locked = False
        .txtcompany.BackColor = vbWhite
        .txtcontactno.Locked = False
        .txtcontactno.BackColor = vbWhite
        .txtcontactnoCode.Locked = False
        .txtcontactnoCode.BackColor = vbWhite
        .txtemail.Locked = False
        .txtemail.BackColor = vbWhite
        .cbo_nationality.Locked = False
        .cbo_nationality.BackColor = vbWhite
        .txthomeadd.Locked = False
        .txthomeadd.BackColor = vbWhite
        .dtpChkOdate.Enabled = True
        .lblsep.BackColor = vbWhite
        .cmdAdd.Enabled = False
        .cmdSave.Enabled = True
    End With
End Sub
   
Public Sub DelReserved()
 cn.Execute "DELETE FROM tblReserved WHERE TransID = " & transID
End Sub
