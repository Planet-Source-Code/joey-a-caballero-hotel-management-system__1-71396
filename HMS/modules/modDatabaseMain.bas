Attribute VB_Name = "modDBMain"
Public cn As New ADODB.Connection
Public lpass As String
Public roomtype As String
Public transID As Integer
Public isPayment As Boolean
Public ischkout As Boolean
Public isRchkin As Boolean
Public isRoomNo As Integer
Public gDays As Integer
Public transIDs As Long
Public isEditInfo As Boolean
Public isRInfo As Boolean
Public rNo As Integer
Public qL As Boolean

    Public sName As String
    Public iRoomNum As Integer
    Public sRoomTar As String
    Public curRoomRate As Currency
    
Dim strServer As String
Dim strDatabase As String
Dim strUser As String
Dim strPass As String

Public Sub dbconnect()
On Error GoTo err:
    '''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''Connection use for Access DB'''''''
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open App.Path & "\database\hms.mdb"
    '''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''
'With cn
'    .ConnectionString = strcon
'    .Open
'End With
    Unload frm_Splash
    frmLogin.Show
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error Connecting Database!"
End Sub


