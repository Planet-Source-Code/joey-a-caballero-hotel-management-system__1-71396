VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmEditUser 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit User"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   StartUpPosition =   2  'CenterScreen
   Begin HMS.b8Line b8Line1 
      Height          =   60
      Left            =   -30
      TabIndex        =   9
      Top             =   2400
      Width           =   9015
      _extentx        =   15901
      _extenty        =   106
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1230
      TabIndex        =   7
      Text            =   "Select Type"
      Top             =   1950
      Width           =   3585
   End
   Begin VB.TextBox txtFullName 
      Height          =   330
      Left            =   1230
      MaxLength       =   60
      TabIndex        =   6
      Top             =   1590
      Width           =   3540
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1230
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1230
      Width           =   3540
   End
   Begin VB.TextBox txtUserName 
      Height          =   330
      Left            =   1230
      MaxLength       =   20
      TabIndex        =   2
      Top             =   810
      Width           =   3540
   End
   Begin HMS.b8SContainer b8SContainer1 
      Height          =   570
      Left            =   -60
      TabIndex        =   0
      Top             =   5520
      Width           =   9135
      _extentx        =   16113
      _extenty        =   1005
      bordercolor     =   14737632
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   405
         Left            =   7500
         TabIndex        =   10
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
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
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   405
         Left            =   5940
         TabIndex        =   11
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
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
   End
   Begin MSComctlLib.ListView listAccess 
      Height          =   2745
      Left            =   60
      TabIndex        =   12
      Top             =   2730
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4842
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "For User"
         Object.Width           =   3175
      EndProperty
   End
   Begin HMS.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   14
      Top             =   510
      Width           =   8985
      _extentx        =   15849
      _extenty        =   106
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit User"
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
      TabIndex        =   15
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restrictions"
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   2490
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   1980
      Width           =   360
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   1590
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   1260
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   840
      Width           =   780
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   0
      Picture         =   "frmEditUser.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8985
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
