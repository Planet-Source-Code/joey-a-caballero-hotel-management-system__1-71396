VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl HSESDataFolder 
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   ScaleHeight     =   5625
   ScaleWidth      =   6960
   Begin MSComctlLib.ImageList imgListEnrolment 
      Left            =   3555
      Top             =   2265
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
            Picture         =   "HSESDataFolder.ctx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFolder 
      Height          =   5550
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   9790
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgListEnrolment"
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
   End
End
Attribute VB_Name = "HSESDataFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Dim slSchoolYearTitle() As String 'holds sy title
Dim slDepartmentTitle() As String 'holds department title
Dim slYearLevelTitle() As String 'holds yl title
Dim slSectionOfferingTitle() As String 'holds SectionOffering

Dim IsStarted As Boolean
'Event Declarations:
Event FolderClick(fNode As Node, sRecordType As String) 'MappingInfo=tvFolder,tvFolder,-1,Click
Event Started()
'----------------------------------------------------------
'START UP
'----------------------------------------------------------
Public Function Start(Optional sSectionOfferingTitle As String = "", Optional sSchoolYearTitle As String = "")
    
    
    'refresh SectionOffering tree
    Refresh_Tree
    
    'set parameter
    If sSectionOfferingTitle <> "" Then
        SetSelectedSectionOffering sSectionOfferingTitle, sSchoolYearTitle
    End If
    IsStarted = True
    RaiseEvent Started
End Function

Public Function Release()
    IsStarted = False
    tvFolder.Nodes.Clear
End Function




Private Function Refresh_Tree()
    'add school year
    Refresh_SchoolYear
    'add Department
    Refresh_Department
    'add year level
    Refresh_YearLevel
    'add SectionOffering
    Refresh_SectionOffering
End Function

Private Function Refresh_SchoolYear()
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    
    'clear tree
    tvFolder.Nodes.Clear
    
    sSQL = "SELECT tblSchoolYear.SchoolYearTitle" & _
            " FROM tblSchoolYear;"
    
    If ConnectRS(HSESDB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
    
    ReDim slSchoolYearTitle(getRecordCount(vRS) - 1)
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        slSchoolYearTitle(i) = ReadField(vRS.Fields("SchoolYearTitle"))
        AddSchoolYearToTree slSchoolYearTitle(i)
        
        
        i = i + 1
        vRS.MoveNext
    Wend
    
    
    
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function Refresh_Department()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    
    
    
    
    sSQL = "SELECT tblDepartment.DepartmentTitle" & _
            " FROM tblDepartment"

    If ConnectRS(HSESDB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slDepartmentTitle(getRecordCount(vRS) - 1)
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
    
        slDepartmentTitle(i) = ReadField(vRS.Fields("DepartmentTitle"))
        
        For ii = 0 To UBound(slSchoolYearTitle)
            AddDepartmentToTree slSchoolYearTitle(ii), slDepartmentTitle(i)
        Next
        
        i = i + 1
        vRS.MoveNext
    Wend
        
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function Refresh_YearLevel()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    Dim iii As Integer
    
    
    
    sSQL = "SELECT tblYearLevel.YearLevelTitle" & _
            " FROM tblYearLevel"

    If ConnectRS(HSESDB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slYearLevelTitle(getRecordCount(vRS) - 1)
    
    i = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        
        slYearLevelTitle(i) = ReadField(vRS.Fields("YearLevelTitle"))
        
        For ii = 0 To UBound(slSchoolYearTitle)
            For iii = 0 To UBound(slDepartmentTitle)
                AddYearLevelToTree slSchoolYearTitle(ii), slDepartmentTitle(iii), slYearLevelTitle(i)
            Next
        Next
        
        i = i + 1
        
        vRS.MoveNext
    Wend
        
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function Refresh_SectionOffering()

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim ii As Integer
    
    
    
    sSQL = "SELECT tblSectionOffering.SectionOfferingID, tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle" & _
            " FROM tblYearLevel INNER JOIN ((tblDepartment INNER JOIN tblSection ON tblDepartment.DepartmentID = tblSection.DepartmentID) INNER JOIN tblSectionOffering ON tblSection.SectionID = tblSectionOffering.SectionID) ON tblYearLevel.YearLevelID = tblSection.YearLevelID" & _
            " GROUP BY tblSectionOffering.SectionOfferingID, tblSectionOffering.SchoolYear, tblDepartment.DepartmentTitle, tblYearLevel.YearLevelTitle, tblSection.SectionTitle;"

    If ConnectRS(HSESDB, vRS, sSQL) <> True Then
        GoTo RealeaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RealeaseAndExit
    End If
        
    ReDim slSectionOfferingTitle(getRecordCount(vRS) - 1)
    
    ii = 0
    vRS.MoveFirst
    
    While vRS.EOF = False
        slSectionOfferingTitle(ii) = ReadField(vRS.Fields("SectionTitle"))
    
        AddSectionOfferingToTree ReadField(vRS.Fields("SchoolYear")), ReadField(vRS.Fields("DepartmentTitle")), ReadField(vRS.Fields("YearLevelTitle")), slSectionOfferingTitle(ii), ReadField(vRS.Fields("SectionOfferingID"))
        
        ii = ii + 1
        vRS.MoveNext
    Wend
        
    
RealeaseAndExit:
    Set vRS = Nothing
End Function


Private Function AddSchoolYearToTree(sSchoolYearTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = keySchoolYear & ";" & sSchoolYearTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add , , keySchoolYear & ";" & sSchoolYearTitle, sSchoolYearTitle, 1
End Function



Private Function AddDepartmentToTree(sSchoolYearTitle As String, sDepartmentTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = KeyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add keySchoolYear & ";" & sSchoolYearTitle, tvwChild, KeyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, sDepartmentTitle, 1
End Function


Private Function AddYearLevelToTree(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = KeyYearLevel & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add KeyDepartment & ";" & sSchoolYearTitle & ";" & sDepartmentTitle, tvwChild, KeyYearLevel & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle, sYearLevelTitle, 1

End Function


Private Function AddSectionOfferingToTree(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String, sSectionOfferingTitle As String, sKey As String)
    Dim tNode As Node
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = KeySectionOffering & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle & ";" & sSectionOfferingTitle Then
            Exit Function
        End If
    Next
    
    tvFolder.Nodes.Add KeyYearLevel & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle, tvwChild, KeySectionOffering & ";" & sSchoolYearTitle & ";" & sDepartmentTitle & ";" & sYearLevelTitle & ";" & sKey, sSectionOfferingTitle, 1

End Function

Private Function SetSelectedSectionOffering(sSectionOfferingTitle As String, Optional sSchoolYearTitle As String = "")
    Dim tNode As Node
    Dim splitKey() As String


    For Each tNode In tvFolder.Nodes
    
        If tNode.Text = sSectionOfferingTitle And Left(tNode.Key, 4) = KeySectionOffering Then
            
            splitKey = Split(tNode.Key, ";")
            
            If sSchoolYearTitle = "" Then
                tNode.Selected = True
                tNode.EnsureVisible
                Exit For
            Else
            
                If splitKey(1) = sSchoolYearTitle Then
                    tNode.Selected = True
                    tNode.EnsureVisible
                    Exit For
                End If
            End If
            
        End If
        
    Next
End Function



Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelY
End Function
Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelX
End Function

Private Sub tvFolder_Click()


    RaiseEvent FolderClick(tvFolder.SelectedItem, Left(tvFolder.SelectedItem.Key, 4))
    
End Sub







Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.ScaleMode = vbPixels
    tvFolder.Move 1, 1, GetWidth - 2, GetHeight - 2
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetSchoolYearList(ByRef sList() As String) As Boolean
    If UBound(slSchoolYearTitle) < 0 Then
        GetSchoolYearList = False
    Else
        sList = slSchoolYearTitle
        GetSchoolYearList = True
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetYearLevelList(ByRef sList() As String) As Variant
    sList = slYearLevelTitle
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetSectionOfferingList(ByRef sList() As String) As Variant
    sList = slSectionOfferingTitle
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetDepartmentList(ByRef sList() As String) As Variant
    sList = slDepartmentTitle
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetSchoolYearChilds(sSchoolYearTitle As String, ByRef sKey() As String, ByRef sText() As String) As Boolean
    Dim tNode As Node
    Dim NodeCount As Integer
    Dim i As Integer
    Dim splitKey() As String
    
    NodeCount = 0
    
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeyDepartment And splitKey(1) = sSchoolYearTitle Then
            NodeCount = NodeCount + 1
        End If
    Next
    
    If NodeCount < 1 Then
        GetSchoolYearChilds = False
        Exit Function
    End If
    
    ReDim sKey(NodeCount - 1)
    ReDim sText(NodeCount - 1)

    i = 0
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeyDepartment And splitKey(1) = sSchoolYearTitle Then
             sKey(i) = tNode.Key
             sText(i) = tNode.Text
            i = i + 1
        End If
    Next
    
    GetSchoolYearChilds = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetDepartmentChilds(sSchoolYearTitle As String, sDepartmentTitle As String, ByRef sKey() As String, ByRef sText() As String) As Boolean
    Dim tNode As Node
    Dim NodeCount As Integer
    Dim i As Integer
    Dim splitKey() As String
    
    NodeCount = 0
    
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeyYearLevel Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle Then
                NodeCount = NodeCount + 1
            End If
        End If
    Next
    
    If NodeCount < 1 Then
        GetDepartmentChilds = False
        Exit Function
    End If
    
    ReDim sKey(NodeCount - 1)
    ReDim sText(NodeCount - 1)
    
    i = 0
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeyYearLevel Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle Then
                sKey(i) = tNode.Key
             sText(i) = tNode.Text
                i = i + 1
            End If
        End If
    Next
    
    GetDepartmentChilds = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetYearLevelChilds(sSchoolYearTitle As String, sDepartmentTitle As String, sYearLevelTitle As String, ByRef sKey() As String, ByRef sText() As String) As Boolean
    Dim tNode As Node
    Dim NodeCount As Integer
    Dim i As Integer
    Dim splitKey() As String
    
    NodeCount = 0
    
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeySectionOffering Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle And splitKey(3) = sYearLevelTitle Then
                NodeCount = NodeCount + 1
            End If
        End If
    Next
    
    If NodeCount < 1 Then
        GetYearLevelChilds = False
        Exit Function
    End If
    
    ReDim sKey(NodeCount - 1)
    ReDim sText(NodeCount - 1)
    
    i = 0
    For Each tNode In tvFolder.Nodes
        splitKey = Split(tNode.Key, ";")
        If splitKey(0) = KeySectionOffering Then
            If splitKey(1) = sSchoolYearTitle And splitKey(2) = sDepartmentTitle And splitKey(3) = sYearLevelTitle Then
                sKey(i) = tNode.Key
             sText(i) = tNode.Text
                i = i + 1
            End If
        End If
    Next
    
    GetYearLevelChilds = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SelectNode(sKey As String) As Variant

    Dim tNode As Node
    
    If IsStarted = False Then
        Start
    End If
    
    For Each tNode In tvFolder.Nodes
        If tNode.Key = sKey Then

            tNode.Selected = True
            tvFolder_Click
        End If
    Next

End Function

