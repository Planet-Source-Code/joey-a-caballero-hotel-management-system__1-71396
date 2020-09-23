VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{78FD1924-C909-4CD4-BC9A-A41DAE171DCE}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hotel Management System"
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11850
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   3840
      Top             =   1200
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":17D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1857E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":421A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6BDC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6C35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":71F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9BBA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A30A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CDD7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CE656
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CEF30
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CF80A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CF996
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":F95B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer timerVote 
      Interval        =   10000
      Left            =   5340
      Top             =   2880
   End
   Begin VB.Timer timerFormTab 
      Interval        =   1
      Left            =   6480
      Top             =   2400
   End
   Begin VB.Timer timerWatchCursor 
      Interval        =   1000
      Left            =   4680
      Top             =   2280
   End
   Begin VB.PictureBox tbMain 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   11850
      TabIndex        =   3
      Top             =   0
      Width           =   11850
      Begin VB.Timer timerUpdateDate 
         Interval        =   1000
         Left            =   3840
         Top             =   600
      End
      Begin VB.PictureBox bgTool 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   2160
         ScaleHeight     =   705
         ScaleWidth      =   6705
         TabIndex        =   7
         Top             =   0
         Width           =   6705
         Begin HMS.b8ToolButton cmdSysLock 
            Height          =   735
            Left            =   240
            TabIndex        =   8
            Top             =   0
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1296
            Picture         =   "mdiMain.frx":1231DA
            Caption         =   "&Lock"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
         End
         Begin HMS.b8ToolButton cmdAbout 
            Height          =   615
            Left            =   3450
            TabIndex        =   9
            Top             =   30
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":12A6DC
            Caption         =   "&About"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
         End
         Begin HMS.b8ToolButton cmdExit 
            Height          =   615
            Left            =   4890
            TabIndex        =   10
            Top             =   30
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":12AFE2
            Caption         =   "&Exit"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
         End
         Begin HMS.b8ToolButton cmdDummy1 
            Height          =   615
            Left            =   5340
            TabIndex        =   11
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   1085
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   12307149
            Enabled         =   0   'False
         End
         Begin HMS.b8ToolButton cmdSecurity 
            Height          =   615
            Left            =   1800
            TabIndex        =   16
            Top             =   30
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":130D44
            Caption         =   "&Security"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            X1              =   30
            X2              =   21700
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            X1              =   330
            X2              =   22000
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            X1              =   330
            X2              =   30
            Y1              =   720
            Y2              =   0
         End
         Begin VB.Image Image2 
            Height          =   720
            Left            =   0
            Picture         =   "mdiMain.frx":131598
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image4 
            Height          =   735
            Left            =   0
            Picture         =   "mdiMain.frx":132D1A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   19995
         End
      End
      Begin VB.PictureBox bgTabBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   3720
         ScaleHeight     =   435
         ScaleWidth      =   19155
         TabIndex        =   4
         Top             =   720
         Width           =   19155
         Begin VB.PictureBox bgTab 
            Appearance      =   0  'Flat
            BackColor       =   &H00D8E9EC&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   0
            ScaleHeight     =   435
            ScaleWidth      =   21525
            TabIndex        =   5
            Top             =   0
            Width           =   21525
            Begin lvButton.lvButtons_H cmdOpenForms 
               Height          =   420
               Index           =   0
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Visible         =   0   'False
               Width           =   3420
               _ExtentX        =   6033
               _ExtentY        =   741
               Caption         =   "Quick Launch"
               CapAlign        =   2
               BackStyle       =   4
               Shape           =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   4210752
               cFHover         =   4210752
               cBhover         =   14737632
               Focus           =   0   'False
               cGradient       =   14737632
               Gradient        =   1
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
            End
         End
      End
      Begin VB.Label lbltime 
         Caption         =   "Label1"
         Height          =   15
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   15
      End
      Begin VB.Label lblDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   750
         TabIndex        =   13
         Top             =   510
         Width           =   180
      End
      Begin VB.Label lblusername 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   750
         TabIndex        =   12
         Top             =   300
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   345
         Left            =   90
         Picture         =   "mdiMain.frx":132E1C
         Stretch         =   -1  'True
         Top             =   765
         Width           =   3540
      End
      Begin VB.Label lblCurrentUserName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sytem User:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   195
         Left            =   750
         TabIndex        =   6
         Top             =   60
         Width           =   1035
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   0
         Picture         =   "mdiMain.frx":133014
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   3735
      ScaleHeight     =   343
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   2
      Top             =   1095
      Width           =   15
   End
   Begin VB.Timer timerMonChild 
      Interval        =   1
      Left            =   3960
      Top             =   2280
   End
   Begin VB.Timer timerWritePreLogOut 
      Interval        =   10000
      Left            =   3960
      Top             =   3000
   End
   Begin MSComctlLib.ImageList imgListOption 
      Left            =   5400
      Top             =   1440
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
            Picture         =   "mdiMain.frx":1334E8
            Key             =   "ListOption"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":133A82
            Key             =   "ChangeFont"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMainDisable 
      Left            =   4800
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilMainNormal 
      Left            =   7800
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":13401C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":134629
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTBList 
      Left            =   6600
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":134C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1354F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":135DD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1366AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":136F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":13785E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":138138
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMainHot 
      Left            =   7200
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":138A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1636EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":18D30E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1B6F30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin HMS.b8SideBar SideBar 
      Align           =   3  'Align Left
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      Top             =   1095
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9075
      BackColor       =   -2147483633
      Begin HMS.b8SideTab b8tListOption 
         Height          =   2505
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   360
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   4419
         Caption         =   "Legend"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16777215
         MaxHeight       =   1560
         Begin MSComctlLib.ListView LVL 
            Height          =   2295
            Left            =   0
            TabIndex        =   18
            Top             =   360
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   4048
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            Icons           =   "ilMainHot"
            SmallIcons      =   "ilMainHot"
            ColHdrIcons     =   "ilMainHot"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "mdiMain.frx":1B780A
            NumItems        =   0
         End
      End
      Begin HMS.b8SideTab b8tListOption 
         Height          =   6000
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   0
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   10583
         Caption         =   "Quick Launch"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16777215
         MaxHeight       =   4250
         Begin MSComctlLib.ListView LVQ 
            Height          =   6360
            Left            =   0
            TabIndex        =   15
            Top             =   360
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   11218
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            Icons           =   "ImageList1"
            SmallIcons      =   "ImageList1"
            ColHdrIcons     =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   16777214
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "mdiMain.frx":1B7B24
            NumItems        =   0
         End
      End
      Begin VB.Image imgSideBarBottom 
         Height          =   3450
         Left            =   0
         Picture         =   "mdiMain.frx":1B7E3E
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   3615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File Menu"
      Begin VB.Menu mnuUser 
         Caption         =   "&Create User"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "&Security"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Log &Off"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuReservation 
      Caption         =   "&Reservation"
      Visible         =   0   'False
      Begin VB.Menu mnuHotelRoom 
         Caption         =   "&Hotel Room"
      End
      Begin VB.Menu mnusp5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunctionHall 
         Caption         =   "&Function Hall"
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuRsetup 
         Caption         =   "&Room Setup"
         Begin VB.Menu mnusRoom 
            Caption         =   "&Rooms"
         End
         Begin VB.Menu sp26 
            Caption         =   "-"
         End
         Begin VB.Menu mnurrSetup 
            Caption         =   "R&oom Rate"
         End
      End
      Begin VB.Menu mnu25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsetup 
         Caption         =   "&Meals Setup"
         Begin VB.Menu mnuMealsEntry 
            Caption         =   "&Meals Entry"
         End
         Begin VB.Menu sp27 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCatEntry 
            Caption         =   "&Category Entry"
         End
      End
      Begin VB.Menu mnu24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSsetup 
         Caption         =   "&Services Setup"
      End
      Begin VB.Menu mnu23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDsetup 
         Caption         =   "&Damage Setup"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "R&eport"
      Begin VB.Menu mnuDaily 
         Caption         =   "&Daily"
         Begin VB.Menu mnuPayment 
            Caption         =   "&Payment"
            Begin VB.Menu mnuPaid 
               Caption         =   "P&aid"
            End
            Begin VB.Menu mnusp2 
               Caption         =   "-"
            End
            Begin VB.Menu mnuUnpaid 
               Caption         =   "&Unpaid"
            End
         End
         Begin VB.Menu mnusp3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOccupiedRooms 
            Caption         =   "&Occupied Rooms"
         End
         Begin VB.Menu mnusp4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCustomers 
            Caption         =   "&Customers"
         End
      End
      Begin VB.Menu mnuWeekly 
         Caption         =   "&Weekly"
      End
      Begin VB.Menu mnuMonthly 
         Caption         =   "&Monthly"
      End
      Begin VB.Menu sp18 
         Caption         =   "-"
      End
      Begin VB.Menu mnudccr 
         Caption         =   "&DCCR"
      End
      Begin VB.Menu sp20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDamage 
         Caption         =   "D&amage Report"
      End
      Begin VB.Menu sp22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRes 
         Caption         =   "&Reservation List"
      End
      Begin VB.Menu sp21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsersLog 
         Caption         =   "User's Log"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Us"
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnupcheckin 
         Caption         =   "Check-in"
      End
      Begin VB.Menu mnusp6 
         Caption         =   "-"
      End
      Begin VB.Menu mnupcheckout 
         Caption         =   "Check-out"
      End
      Begin VB.Menu mnupFChkOut 
         Caption         =   "&Force Check-Out"
      End
      Begin VB.Menu mnusp7 
         Caption         =   "-"
      End
      Begin VB.Menu mnusp8 
         Caption         =   "-"
      End
      Begin VB.Menu mnupreservation 
         Caption         =   "Reservation"
         Begin VB.Menu mnupreserve 
            Caption         =   "Reserve"
         End
         Begin VB.Menu mnusp9 
            Caption         =   "-"
         End
         Begin VB.Menu mnupcancel 
            Caption         =   "Cancel"
         End
      End
      Begin VB.Menu sp11 
         Caption         =   "-"
      End
      Begin VB.Menu mnucPayment 
         Caption         =   "Payment"
      End
      Begin VB.Menu sp12 
         Caption         =   "-"
      End
      Begin VB.Menu sp13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddOns 
         Caption         =   "Add Ons"
         Begin VB.Menu mnuMeals 
            Caption         =   "Meals"
         End
         Begin VB.Menu sp14 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddServ 
            Caption         =   "Additonal Services"
         End
      End
      Begin VB.Menu sp15 
         Caption         =   "-"
      End
      Begin VB.Menu sp16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDamages 
         Caption         =   "Damages"
      End
      Begin VB.Menu sp17 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnupopEdit 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEditCustoInfo 
         Caption         =   "&Update Customer's Information"
      End
   End
   Begin VB.Menu mnulPay 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnulPayment 
         Caption         =   "Payment"
      End
   End
   Begin VB.Menu mnupopAvailable 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAvailable 
         Caption         =   "&Available"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastSideTabOnFocus As Integer

Dim defCmdOpenFormsTop As Integer
Dim defCmdOpenFormsLeft As Integer
Dim defCmdOpenFormsWidth As Integer

Dim VisibleTabCount As Integer

Dim fn(100) As String
Dim rs As New ADODB.Recordset


Private Sub b8tListOption_CompleteContract(Index As Integer)
If b8tListOption(0).Top < 0 Then
        SideBar.MoveDownControls (0 - b8tListOption(0).Top)
    End If
    
    Select Case Index
        Case 0 'info
             b8tListOption(Index).ForeColor = b8tListOption(Index).ContractedForeColor
    End Select
End Sub

Private Sub b8tListOption_Resize(Index As Integer)
Dim i As Integer
Dim iSpaceExceed As Integer
  
    For i = 1 To b8tListOption.UBound
        b8tListOption(i).Top = b8tListOption(i - 1).Top + b8tListOption(i - 1).Height - Screen.TwipsPerPixelY
    Next
        
    If Index > 0 Then
        b8tListOption(Index).Top = b8tListOption(Index - 1).Top + b8tListOption(Index - 1).Height - Screen.TwipsPerPixelY
    End If
    
       iSpaceExceed = (b8tListOption(Index).Top + b8tListOption(Index).Height) - SideBar.Height

      If iSpaceExceed > 0 Then
          If iSpaceExceed - b8tListOption(Index).Top > 0 Then
              iSpaceExceed = b8tListOption(Index).Top
           End If

            SideBar.MoveUpControls iSpaceExceed
        End If

    SideBar.CheckExceedControl

    'If b8tListOption(Index).AutoExpand = True Then
    '    For i = 0 To b8tListOption.UBound
    '        If Index <> i And b8tListOption(i).AutoExpand = True Then
    '            b8tListOption(i).Expanded = False
    '        End If
    '    Next
    'End If
End Sub

Private Sub cmdAbout_Click()
    MsgBox "Please Code on it"
End Sub
Private Sub cmdExit_Click()
    If MsgBox("You are about to close the system." & vbCrLf & vbCrLf & "Are you sure?", vbQuestion + vbYesNo, "Exit Confirmation") = vbYes Then
        End
    End If
End Sub

Private Function RefreshTabButtonFace(Index As Integer)
    Dim i As Integer
    
    cmdOpenForms(Index).ButtonStyle = lv_Flat
    Set cmdOpenForms(Index).Picture = Me.ActiveForm.Icon
    cmdOpenForms(Index).BackColor = &HFFFFFF
    cmdOpenForms(Index).CaptionAlign = vbLeftJustify
    cmdOpenForms(Index).FontStyle = lv_Bold
    
    
    For i = 0 To cmdOpenForms.UBound
        If Index <> i Then
            cmdOpenForms(i).FontStyle = lv_PlainStyle
            cmdOpenForms(i).BackColor = &HD8E9EC
            cmdOpenForms(Index).CaptionAlign = vbLeftJustify
            cmdOpenForms(i).ButtonStyle = lv_hover
        End If
    Next
    
End Function

Public Function LockApp()
    frmLock.ShowForm
End Function

Private Sub cmdOpenForms_Click(Index As Integer)
    On Error Resume Next
    Dim Frm As Form
    
    

    For Each Frm In Forms
        If Frm.Name = fn(Index) Then
            If Me.ActiveForm.Name <> fn(Index) Then
            Frm.SetFocus
            End If
            Exit For
        End If
    Next
    
    RefreshTabButtonFace Index
End Sub

Private Sub cmdSysLock_Click()
    frmLock.ShowForm
End Sub


Private Sub LVQ_DblClick()
    Select Case Me.LVQ.SelectedItem.Key
        Case "Cuser"
            mnuUser_Click
        Case "SLock"
            cmdSysLock_Click
        Case "uLog"
            mnuUsersLog_Click
        Case "AVacant"
            frm_Welcome.VloadRooms
        Case "AOccupied"
            frm_Welcome.OloadRooms
        Case "AReserved"
        
        Case "AHome"
            frm_Welcome.loadRooms
        Case "About"
        Case "LOff"
            mnulogoff_Click
        Case "ESystem"
            cmdExit_Click
    End Select
End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    SideBar_Resize
    Me.ActiveForm.Form_Refresh
End Sub

Public Function RefreshActiveForm()
    On Error Resume Next
    Me.ActiveForm.Form_Refresh
End Function

Private Sub MDIForm_Load()
    Dim i As Integer
    
    'defaults
    defCmdOpenFormsTop = cmdOpenForms(0).Top
    defCmdOpenFormsLeft = cmdOpenForms(0).Left
    defCmdOpenFormsWidth = cmdOpenForms(0).Width
    
    
    frm_Welcome.Show
    
    'set sidebar
    For i = b8tListOption.UBound To 0 Step -1
        b8tListOption(i).Expanded = False
        b8tListOption(i).ZOrder 0
    Next
    
    loadLegend
    loadsidebar
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    bgTool.Left = mdiMain.Width - bgTool.Width
End Sub

Private Sub mnuFormMenu_Click(Index As Integer)
    On Error Resume Next
    
    'Me.ActiveForm.Form_MenuClick mnuFormMenu(Index).Caption
    
End Sub

Private Sub mnuAddServ_Click()
    With frm_Services
        .lblFloor.Caption = frm_Welcome.lvr.SelectedItem.SubItems(2)
        .lblroomno.Caption = frm_Welcome.lvr.SelectedItem.SubItems(4)
        .lblRT.Caption = frm_Welcome.lvr.SelectedItem.SubItems(3)
        'isPayment = False
        .Show 1
    End With
End Sub

Private Sub mnuAvailable_Click()
    rNo = frm_Welcome.lvr.SelectedItem.SubItems(4)
    pAvailable
End Sub

Private Sub mnuCatEntry_Click()
    frm_Add_Cat.Show 1
End Sub

Private Sub mnucPayment_Click()
    isPayment = True
    ischkout = False
    On Error Resume Next
    transID = frm_Welcome.lvr.SelectedItem.SubItems(5)
    'MsgBox transID
    frm_Payment.Show 1
End Sub


Private Sub mnuCustomers_Click()
    frm_CInfo.Show
End Sub

Private Sub mnuDamage_Click()
    frm_Damage_Report.Show
End Sub

Private Sub mnuDamages_Click()
    With frm_Damages
        .lblFloor.Caption = frm_Welcome.lvr.SelectedItem.SubItems(2)
        .lblroomno.Caption = frm_Welcome.lvr.SelectedItem.SubItems(4)
        .lblRT.Caption = frm_Welcome.lvr.SelectedItem.SubItems(3)
        'isPayment = False
        .Show 1
    End With
End Sub

Private Sub mnudccr_Click()
    frm_DCCR.Show
End Sub

Private Sub mnuDsetup_Click()
    frm_Add_Damages.Show 1
End Sub

Private Sub mnuEditCustoInfo_Click()
    frm_Update_Info.Show 1
End Sub

Private Sub mnulogoff_Click()
    If MsgBox("You are about to Log Off." & vbCrLf & vbCrLf & "Are you sure?", vbQuestion + vbYesNo, "Log Off Confirmation") = vbYes Then
        Unload Me
        frmLogin.Show
    End If
End Sub

Private Sub mnulPayment_Click()
    With frm_Late_Payment
        .txtAmountDue.Text = frm_Unpaid.listRecord.SelectedItem.SubItems(10)
        .txtBalance.Text = frm_Unpaid.listRecord.SelectedItem.SubItems(12)
        .txtRoomRate.Text = frm_Unpaid.listRecord.SelectedItem.SubItems(8)
        .lblNoDays.Caption = frm_Unpaid.listRecord.SelectedItem.SubItems(14)
        .lblDate.Caption = frm_Unpaid.listRecord.SelectedItem.SubItems(4)
        .lbltime.Caption = frm_Unpaid.listRecord.SelectedItem.SubItems(5)
        .lblAd.Caption = frm_Unpaid.listRecord.SelectedItem.SubItems(10)
        .lblPaid.Caption = frm_Unpaid.listRecord.SelectedItem.SubItems(11)
        .lblBal.Caption = frm_Unpaid.listRecord.SelectedItem.SubItems(12)
    End With
    transID = frm_Unpaid.listRecord.SelectedItem.SubItems(15)
    frm_Late_Payment.Show 1
End Sub

Private Sub mnuMeals_Click()
    With frm_meals
        .lblFloor.Caption = frm_Welcome.lvr.SelectedItem.SubItems(2)
        .lblroomno.Caption = frm_Welcome.lvr.SelectedItem.SubItems(4)
        .lblRT.Caption = frm_Welcome.lvr.SelectedItem.SubItems(3)
        'isPayment = False
        .Show 1
    End With
End Sub

Private Sub mnuMealsEntry_Click()
    frm_Add_Meals.Show 1
End Sub

Private Sub mnuMonthly_Click()
    frm_Monthly_Report.Show
End Sub

Private Sub mnuOccupiedRooms_Click()
    frm_Daily_Report.Show
End Sub


Private Sub mnuPaid_Click()
    frm_Paid.Show
End Sub

Private Sub mnupcancel_Click()
On Error GoTo err:
    If MsgBox("You are about to Cancel Reservation" & vbCrLf & vbCrLf & "Are you sure", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblReserved where TransID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        rs.Delete
        
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblCustomerInfo where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        rs.Delete
        
        MsgBox "Reservation Successfully Canceled", vbInformation
    End If
Exit Sub
err:
MsgBox err.Description, vbCritical
End Sub

Private Sub mnupcheckin_Click()
    isPayment = False
    frm_Welcome.chkin
End Sub

Private Sub mnupcheckout_Click()
On Error GoTo err:
    Dim bal, paid As Currency
    Dim rRate As Currency
    ischkout = True
    transID = frm_Welcome.lvr.SelectedItem.SubItems(5)
    rNo = frm_Welcome.lvr.SelectedItem.SubItems(4)
    
    If rs.State = adStateOpen Then rs.Close
        'rs.Open "SELECT tblPayment.RoomRate, Max(tblPayment.NoDays) AS MaxOfNoDays, Sum(tblPayment.cPaid) AS SumOfcPaid, Max(tblPayment.AmountDue) AS MaxOfAmountDue, [MaxOfAmountDue]-[SumOfcPaid] AS Expr1, tblPayment.TransactionID " & _
        "From tblPayment Where (((tblPayment.TransactionID) = " & transID & "))" & _
        " GROUP BY tblPayment.RoomRate, tblPayment.TransactionID;", cn, adOpenKeyset, adLockPessimistic
    
        'rs.Open "SELECT tblPayment.RoomRate, Max(tblPayment.NoDays) AS MaxOfNoDays, Sum(tblPayment.cPaid) AS SumOfcPaid, Max(tblPayment.AmountDue) AS MaxOfAmountDue, [MaxOfAmountDue]-[SumOfcPaid] AS Expr1, tblPayment.TransactionID " & _
        "From tblPayment Where (((tblPayment.TransactionID) = " & transID & "))" & _
        " GROUP BY tblPayment.RoomRate, tblPayment.TransactionID;", cn, adOpenKeyset, adLockPessimistic
        
       rs.Open "SELECT tblPayment.RoomRate, Max(tblPayment.NoDays) AS MaxOfNoDays, Sum(tblPayment.cPaid) AS SumOfcPaid, Max(tblPayment.AmountDue) AS MaxOfAmountDue, tblPayment.TransactionID, Sum(tblPayment.Damages) AS SumOfDamages, Sum(tblPayment.Services) AS SumOfServices, Sum(tblPayment.Meals) AS SumOfMeals, Max(tblPayment.AmountDue)+Sum(tblPayment.Damages)+Sum(tblPayment.Services)+Sum(tblPayment.Meals)-Sum(tblPayment.cPaid) AS Expr1 " & _
        "From tblPayment Where (((tblPayment.TransactionID) =" & transID & "))GROUP BY tblPayment.RoomRate, tblPayment.TransactionID;", cn, adOpenKeyset, adLockPessimistic
    
        'rs.Open "Select * from tblpayment where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        paid = rs.Fields("SumOfcPaid").Value
        'MsgBox paid
        'MsgBox paid
        rRate = rs.Fields("RoomRate").Value
        'MsgBox rRate
        'getNoDays
        'MsgBox rRate * gDays
        bal = rs.Fields("Expr1").Value 'rRate * gDays - paid
        'MsgBox bal
        If bal > 0 Then
            MsgBox "Occupant of this room still has a balance. Please ask for" & vbCrLf & "Full Payment before check-out" & vbCrLf & vbCrLf & "Thank You!", vbInformation
            isPayment = True
            frm_Payment.Show
        ElseIf bal <= 0 Then
            isPayment = False
            pcheckout
        End If
        
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub


Private Sub mnupFChkOut_Click()
    rNo = frm_Welcome.lvr.SelectedItem.SubItems(4)
    Fpcheckout
End Sub

Private Sub mnupreserve_Click()
    frm_Welcome.chkReserved
    frm_Reserved.Show
    mdiMain.Enabled = False
End Sub

Private Sub mnuRes_Click()
    frm_Reserved_info.Show
End Sub

Private Sub mnurrSetup_Click()
    frm_Add_Room_Tariff.Show 1
End Sub

Private Sub mnuSecurity_Click()
    cmdSysLock_Click
End Sub

Private Sub mnusRoom_Click()
    frm_Add_Rooms.Show 1
End Sub

Private Sub mnuSsetup_Click()
    frm_Add_Services.Show 1
End Sub

Private Sub mnuUnpaid_Click()
    frm_Unpaid.Show
End Sub

Private Sub mnuUser_Click()
    frmAddUser.Show 1
End Sub

Private Sub mnuUsersLog_Click()
    frm_Logs.Show
End Sub

Private Sub mnuWeekly_Click()
    frm_Weekly_Report.Show
End Sub


Private Sub SideBar_CtlsScroll()
    imgSideBarBottom.Top = SideBar.Height - imgSideBarBottom.Height
End Sub

Private Sub SideBar_Resize()
    Dim iSpaceExceed As Integer
    
    iSpaceExceed = (b8tListOption(LastSideTabOnFocus).Top + b8tListOption(LastSideTabOnFocus).Height) - SideBar.Height

        If iSpaceExceed > 0 Then
            If iSpaceExceed - b8tListOption(LastSideTabOnFocus).Top > 0 Then
                iSpaceExceed = b8tListOption(LastSideTabOnFocus).Top
            End If

            SideBar.MoveUpControls iSpaceExceed
        Else
        
            iSpaceExceed = SideBar.Height - (b8tListOption(b8tListOption.UBound).Top + b8tListOption(b8tListOption.UBound).Height)

            If iSpaceExceed > 0 And b8tListOption(0).Top < 0 Then
                If b8tListOption(0).Top + iSpaceExceed > 0 Then
                    iSpaceExceed = 0 - b8tListOption(0).Top
                End If
                SideBar.MoveDownControls iSpaceExceed
            End If
        End If
        
    On Error Resume Next
    
    imgSideBarBottom.Top = SideBar.Height - imgSideBarBottom.Height
End Sub

Public Function RegMDIChild(ByRef AForm As Form)
    Dim i As Integer

    
On Error Resume Next

    Refresh_FormTabButtons
    

    For i = 0 To cmdOpenForms.UBound
        If fn(i) = AForm.Name Then
            RefreshTabButtonFace i
            Exit For
        End If
    Next
    
End Function

'Private Function Aform_SetFormMenu()
    
'    Dim sMenu() As String
'    Dim bHasMenu As Boolean
'    Dim i As Integer
    
'    On Error Resume Next
    
    'default
'    bHasMenu = False
    
    'hide old menus
'    For i = 0 To mnuFormMenu.UBound
'        mnuFormMenu(i).Visible = False
'        Unload mnuFormMenu(i)
'    Next
'    mnuSeparatorEdit3.Visible = False
    
    
'    bHasMenu = Me.ActiveForm.Form_GetMenu(sMenu)
    
'    If bHasMenu = False Then
'        Exit Function
'    End If
    
'    mnuSeparatorEdit3.Visible = True
    
'    For i = 0 To UBound(sMenu)
'        Load mnuFormMenu(i)
'        mnuFormMenu(i).Caption = sMenu(i)
'        mnuFormMenu(i).Visible = True
'   Next
    

'End Function

Private Function Refresh_FormTabButtons()
    Static sfn As String
    Dim i As Integer
    Dim x As Integer

    Dim Frm As Form
    Dim lv As lvButtons_H
    Dim tLeft As Integer
    Dim tWidth As Integer
    
    On Error Resume Next
    bgTab.Width = mdiMain.Width - bgTabBack.Left

    i = 0
    For Each Frm In Forms
        If LCase(Trim(Frm.Name)) <> LCase(Trim(mdiMain.Name)) Then
        If Frm.MDIChild = True Then
            
            Load cmdOpenForms(i)
            cmdOpenForms(i).Caption = Frm.Caption
            cmdOpenForms(i).Visible = True
            fn(i) = Frm.Name
            
            i = i + 1
        End If
        End If
    Next
    

    While i <= cmdOpenForms.UBound
        cmdOpenForms(i).Visible = False
        Unload cmdOpenForms(i)
        i = i + 1
    Wend
    
    tLeft = 0
    tWidth = IIf((bgTab.Width / i) < defCmdOpenFormsWidth, (bgTab.Width / i), defCmdOpenFormsWidth)
    
    For i = 0 To cmdOpenForms.UBound
        If cmdOpenForms(i).Visible = True Then
        
            cmdOpenForms(i).Left = tLeft
            cmdOpenForms(i).Width = tWidth
            tLeft = cmdOpenForms(i).Left + cmdOpenForms(i).Width
            
        End If
    Next
    

    If sfn <> Me.ActiveForm.Name Then
        For i = 0 To cmdOpenForms.UBound
        If fn(i) = Me.ActiveForm.Name Then
            RefreshTabButtonFace i
            Exit For
        End If
    Next
    
    End If
    sfn = Me.ActiveForm.Name
End Function

Private Sub timerFormTab_Timer()
    Refresh_FormTabButtons
End Sub

Private Sub timerMonChild_Timer()
    On Error GoTo ErrShowWelcomeScreen
    Dim s As String
    s = Me.ActiveForm.Name
    
    Exit Sub
    
ErrShowWelcomeScreen:
    frm_Welcome.Show
    timerMonChild.Enabled = False
End Sub

Private Sub timerUpdateDate_Timer()
    lblDate.Caption = "Today is: " & FormatDateTime(Now, vbGeneralDate)
    lbltime.Caption = FormatDateTime(Now, vbLongTime)
    If Me.lbltime.Caption = "12:00:00 AM" Then frm_Welcome.loadRooms
    'frm_Welcome.loadRooms
End Sub

Private Sub timerWatchCursor_Timer()

    Static ic As Integer
    Static op As POINTAPI
    
    Dim p As POINTAPI
    
    GetCursorPos p
    If (p.x < (op.x + 5) And p.x > op.x - 5) And (p.Y < (op.Y + 5) And p.Y > op.Y - 5) Then
        ic = ic + 1
    Else
        ic = 0
    End If
    
    op.x = p.x
    op.Y = p.Y

    If ic > AppSet_LockTimeOut Then
        ic = 0
        Call LockApp
    End If
End Sub

Public Sub loadQuickLauch()
    'Me.LVQ.ListItems.Add , , "Check-in", 2
    'Me.LVQ.ListItems.Add , , "Check-out", 3
        Me.LVQ.ListItems.Add , "Cuser", "Create User", 1
        Me.LVQ.ListItems.Add , "uLog", "User's Log", 11
        Me.LVQ.ListItems.Add , "SLock", "System Lock", 7
    
        Me.LVQ.ListItems.Add , "AVacant", "All Vacant", 6
        Me.LVQ.ListItems.Add , "AOccupied", "All Occupied", 8
        Me.LVQ.ListItems.Add , "AReserved", "All Reserved", 14
        Me.LVQ.ListItems.Add , "AHome", "Home", 9
        
        Me.LVQ.ListItems.Add , "About", "About", 4
        Me.LVQ.ListItems.Add , "LOff", "Log Off", 12
        Me.LVQ.ListItems.Add , "ESystem", "Exit System", 13
        
    
End Sub

Private Sub loadLegend()
    Me.LVL.ListItems.Add , , "Available", 3
    Me.LVL.ListItems.Add , , "Occupied", 1
    Me.LVL.ListItems.Add , , "Reserved", 2
End Sub

Private Sub loadsidebar()
    b8tListOption(0).Expanded = True
    b8tListOption(1).Expanded = True
End Sub

Public Sub pcheckout()
On Error GoTo err:
If MsgBox("You are about to Check-Out Occupant of the Room " & rNo & "." & vbCrLf & vbCrLf & "Are You Sure?", vbQuestion + vbYesNo, "Check-Out Confirmation") = vbYes Then
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblroom where Room_no = " & rNo & ";", cn, adOpenKeyset, adLockPessimistic
    rs.Fields("occupied").Value = "Dirty"
    'rs.Fields("TransID").Value = 0
    rs.Update
    frm_Welcome.loadRooms
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblCustomerInfo where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
    rs.Fields("Stat").Value = "C/OUT"
    rs.Fields("coDate").Value = FormatDateTime(Now, vbShortDate)
    rs.Fields("coTime").Value = FormatDateTime(Now, vbShortTime)
    rs.Update
End If
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub
'Private Sub getNoDays()
'    Dim dtnow As Date
'    Dim dtin As Date
'    Dim stime As Double
    
'    dtnow = FormatDateTime(Now, vbShortDate)
'
'        If rs.State = adStateOpen Then rs.Close
'        rs.Open "Select * from tblCustomerInfo where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
'        dtin = rs.Fields("occu_date").Value
'        stime = Format(rs.Fields("occu_time").Value, "HHMMSS")
'        'MsgBox "hahay"
'    ' Check-in form 12:00am - 5:00am
'    ' Code Ok
'    ' For Adjustmet
'    If stime >= 100 And stime <= 50000 Then
'
'        If DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
'            gDays = DateDiff("d", dtin, dtnow) + 1.5
'        ElseIf DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 140000 Then
'            gDays = DateDiff("d", dtin, dtnow) + 2
'        Else
'            gDays = DateDiff("d", dtin, dtnow) + 1
'        End If
'    ' Check-in form 5:01am - 7:00am
'    ' Code Ok
'    ' For Adjustmet
'    ElseIf stime >= 50100 And stime <= 70000 Then
'    'MsgBox "Brrrrrrrrr   hhhhhhhhhhhhhhhhh"
'        If DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
'            gDays = DateDiff("d", dtin, dtnow) + 1
'        ElseIf DateDiff("d", dtin, dtnow) >= 0 And Format(Time, "HHMMSS") >= 140000 Then
'            gDays = DateDiff("d", dtin, dtnow) + 1.5
'        Else
'            gDays = DateDiff("d", dtin, dtnow) + 0.5
'        End If
'    ' Will Follow
'    ' Code Ok
'    ' For Adjustmet
'    ElseIf DateDiff("d", dtin, dtnow) = 0 And stime >= 70000 And stime <= 120000 Then
'    'MsgBox "Brrrrrrrrr"
'        If DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") <= 140000 Then
'            gDays = DateDiff("d", dtin, dtnow) + 0.5
'        ElseIf DateDiff("d", dtin, dtnow) = 0 And Format(Time, "HHMMSS") >= 140100 Then
'            gDays = DateDiff("d", dtin, dtnow) + 1
'        End If
'    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") >= 120100 And Format(Time, "HHMMSS") <= 140000 Then
'    'MsgBox "B"
'        gDays = DateDiff("d", dtin, dtnow) + 0.5
'    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") >= 140100 Then
'    'MsgBox "R"
'        gDays = DateDiff("d", dtin, dtnow) + 1
'    ElseIf DateDiff("d", dtin, dtnow) <> 0 And Format(Time, "HHMMSS") <= 120100 Then
'    'MsgBox "ambot"
'        gDays = DateDiff("d", dtin, dtnow)
'    Else
'        gDays = 1
'    End If
'        'gDays
'End Sub

Public Sub Fpcheckout()
On Error GoTo err:
If MsgBox("You are about to Check-Out Occupant of the Room " & rNo & "." & vbCrLf & vbCrLf & "Are You Sure?", vbQuestion + vbYesNo, "Check-Out Confirmation") = vbYes Then
    If MsgBox("This is a Force Checkout!" & vbCrLf & vbCrLf & "Are You Sure?", vbQuestion + vbYesNo, "Check-Out Confirmation") = vbYes Then
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblroom where Room_no = " & rNo & ";", cn, adOpenKeyset, adLockPessimistic
        rs.Fields("occupied").Value = "Dirty"
        'rs.Fields("TransID").Value = 0
        rs.Update
        frm_Welcome.loadRooms
        
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblCustomerInfo where TransactionID = " & transID & ";", cn, adOpenKeyset, adLockPessimistic
        rs.Fields("Stat").Value = "C/OUT"
        rs.Fields("coDate").Value = FormatDateTime(Now, vbShortDate)
        rs.Fields("coTime").Value = FormatDateTime(Now, vbShortTime)
        rs.Update
    End If
Exit Sub
End If
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Public Sub loadQuickLauchC()
    'Me.LVQ.ListItems.Add , , "Check-in", 2
    'Me.LVQ.ListItems.Add , , "Check-out", 3
        
        Me.LVQ.ListItems.Add , "uLog", "User's Log", 11
        Me.LVQ.ListItems.Add , "SLock", "System Lock", 7
    
        Me.LVQ.ListItems.Add , "AVacant", "All Vacant", 6
        Me.LVQ.ListItems.Add , "AOccupied", "All Occupied", 8
        Me.LVQ.ListItems.Add , "AReserved", "All Reserved", 14
        Me.LVQ.ListItems.Add , "AHome", "Home", 9
        
        Me.LVQ.ListItems.Add , "About", "About", 4
        Me.LVQ.ListItems.Add , "LOff", "Log Off", 12
        Me.LVQ.ListItems.Add , "ESystem", "Exit System", 13
    
End Sub

Public Sub pAvailable()
On Error GoTo err:
If MsgBox("You are about to make this room Available." & vbCrLf & vbCrLf & "Are You Sure?", vbQuestion + vbYesNo) = vbYes Then
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblroom where Room_no = " & rNo & ";", cn, adOpenKeyset, adLockPessimistic
    rs.Fields("occupied").Value = "false"
    rs.Fields("TransID").Value = 0
    rs.Update
    frm_Welcome.loadRooms
End If
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub
