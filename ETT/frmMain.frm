VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000E&
   Caption         =   "Main Control Panel Of AIR TICKTING System "
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Timer Help_Timer2 
      Left            =   12600
      Top             =   9720
   End
   Begin VB.Timer Help_Timer1 
      Left            =   12600
      Top             =   9240
   End
   Begin VB.Timer Secur_Timer2 
      Left            =   12600
      Top             =   8280
   End
   Begin VB.Timer Secur_Timer1 
      Left            =   12600
      Top             =   7800
   End
   Begin VB.Timer Rpt_Timer2 
      Left            =   12600
      Top             =   6120
   End
   Begin VB.Timer Rpt_Timer1 
      Left            =   12600
      Top             =   5640
   End
   Begin VB.Timer Nav_Timer1 
      Left            =   12600
      Top             =   3600
   End
   Begin VB.Timer Nav_Timer2 
      Left            =   12600
      Top             =   4080
   End
   Begin VB.Timer Admin_Timer1 
      Left            =   12600
      Top             =   1560
   End
   Begin VB.Timer Admin_Timer2 
      Left            =   12600
      Top             =   2040
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10920
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A12A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ACFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B60B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BE34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C743
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E44D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E767
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB05
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EEA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1088C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12586
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14140
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F92
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1599A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox TIMELOG_PIC 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":16A2C
      ScaleHeight     =   495
      ScaleWidth      =   15240
      TabIndex        =   2
      Top             =   0
      Width           =   15240
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblTime"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   9600
         TabIndex        =   32
         Top             =   105
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&&"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   9360
         TabIndex        =   31
         Top             =   90
         Width           =   135
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblDate"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   7440
         TabIndex        =   30
         Top             =   105
         Width           =   1800
      End
      Begin VB.Label lblUserID 
         BackStyle       =   0  'Transparent
         Caption         =   "lblUserID"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   2040
         TabIndex        =   29
         Top             =   105
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Log In Date && Time :  "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   270
         Left            =   5340
         TabIndex        =   5
         Top             =   90
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name : "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   720
         TabIndex        =   4
         Top             =   90
         Width           =   1185
      End
      Begin VB.Image Img_Time 
         Height          =   360
         Left            =   4920
         Picture         =   "frmMain.frx":19E5D
         Stretch         =   -1  'True
         Top             =   45
         Width           =   360
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   9660
      Left            =   13020
      TabIndex        =   1
      Top             =   495
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   17039
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   9615
         Left            =   0
         Picture         =   "frmMain.frx":1A855
         ScaleHeight     =   9615
         ScaleWidth      =   2235
         TabIndex        =   6
         Top             =   480
         Width           =   2235
         Begin VB.Frame Secur_Frame 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   740
            Left            =   133
            TabIndex        =   26
            Top             =   6960
            Width           =   1935
            Begin VB.Image Image25 
               Height          =   300
               Left            =   60
               Picture         =   "frmMain.frx":1C153
               Stretch         =   -1  'True
               Top             =   387
               Width           =   300
            End
            Begin VB.Image Image24 
               Height          =   300
               Left            =   60
               Picture         =   "frmMain.frx":1E8F5
               Stretch         =   -1  'True
               Top             =   67
               Width           =   300
            End
            Begin VB.Label lblChangePwd 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Change Password"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   400
               MouseIcon       =   "frmMain.frx":1F737
               MousePointer    =   99  'Custom
               TabIndex        =   28
               Top             =   440
               Width           =   1520
            End
            Begin VB.Label lblNewUser 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Create New User"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   400
               MouseIcon       =   "frmMain.frx":1FA41
               MousePointer    =   99  'Custom
               TabIndex        =   27
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.Frame Rpt_Frame 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1150
            Left            =   133
            TabIndex        =   24
            Top             =   5170
            Width           =   1935
            Begin VB.Label lblDaily_Month_Rpt 
               BackStyle       =   0  'Transparent
               Caption         =   "Daily/Monthly"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   480
               MouseIcon       =   "frmMain.frx":1FD4B
               MousePointer    =   99  'Custom
               TabIndex        =   35
               Top             =   810
               Width           =   1425
            End
            Begin VB.Label lblDoneJob_Rpt 
               BackStyle       =   0  'Transparent
               Caption         =   "Done Jobs"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   480
               MouseIcon       =   "frmMain.frx":20055
               MousePointer    =   99  'Custom
               TabIndex        =   34
               Top             =   423
               Width           =   1420
            End
            Begin VB.Label lblJobpending_Rpt 
               BackStyle       =   0  'Transparent
               Caption         =   "Pending Jobs"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   480
               MouseIcon       =   "frmMain.frx":2035F
               MousePointer    =   99  'Custom
               TabIndex        =   33
               Top             =   120
               Width           =   1420
            End
            Begin VB.Image Image21 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":20669
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":20973
               Stretch         =   -1  'True
               Top             =   780
               Width           =   300
            End
            Begin VB.Image Image20 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":2420B
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":24515
               Stretch         =   -1  'True
               Top             =   400
               Width           =   300
            End
            Begin VB.Image Image19 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":27EB1
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":281BB
               Stretch         =   -1  'True
               Top             =   80
               Width           =   300
            End
         End
         Begin VB.Frame Nav_Frame 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   133
            TabIndex        =   15
            Top             =   2540
            Width           =   1935
            Begin VB.Image Image3 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":2BA91
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":2BD9B
               Stretch         =   -1  'True
               Top             =   395
               Width           =   300
            End
            Begin VB.Label lblCheckStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "Check Status"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   525
               MouseIcon       =   "frmMain.frx":2F738
               MousePointer    =   99  'Custom
               TabIndex        =   37
               Top             =   440
               Width           =   1365
            End
            Begin VB.Image Image1 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":2FA42
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":2FD4C
               Stretch         =   -1  'True
               Top             =   75
               Width           =   300
            End
            Begin VB.Label lblNewJob 
               BackStyle       =   0  'Transparent
               Caption         =   "New Job"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   525
               MouseIcon       =   "frmMain.frx":3064B
               MousePointer    =   99  'Custom
               TabIndex        =   36
               Top             =   105
               Width           =   1365
            End
            Begin VB.Image Image15 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":30955
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":30C5F
               Stretch         =   -1  'True
               Top             =   1395
               Width           =   300
            End
            Begin VB.Image Image14 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":347F8
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":34B02
               Stretch         =   -1  'True
               Top             =   1740
               Width           =   300
            End
            Begin VB.Image Image13 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":3834F
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":38659
               Stretch         =   -1  'True
               Top             =   1080
               Width           =   300
            End
            Begin VB.Image Image11 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":3C202
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":3C50C
               Stretch         =   -1  'True
               Top             =   750
               Width           =   300
            End
            Begin VB.Label lblCharges 
               BackStyle       =   0  'Transparent
               Caption         =   "Char&ges Apply"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   525
               MouseIcon       =   "frmMain.frx":3FCFF
               MousePointer    =   99  'Custom
               TabIndex        =   19
               Top             =   1755
               Width           =   1365
            End
            Begin VB.Label lblServices 
               BackStyle       =   0  'Transparent
               Caption         =   "Ser&vices Info"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   525
               MouseIcon       =   "frmMain.frx":40009
               MousePointer    =   99  'Custom
               TabIndex        =   18
               Top             =   1425
               Width           =   1365
            End
            Begin VB.Label lblDestination 
               BackStyle       =   0  'Transparent
               Caption         =   "D&estination Info"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   525
               MouseIcon       =   "frmMain.frx":40313
               MousePointer    =   99  'Custom
               TabIndex        =   17
               Top             =   1110
               Width           =   1365
            End
            Begin VB.Label lblAirLine 
               BackStyle       =   0  'Transparent
               Caption         =   "A&ir Line Info"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   525
               MouseIcon       =   "frmMain.frx":4061D
               MousePointer    =   99  'Custom
               TabIndex        =   16
               Top             =   780
               Width           =   1365
            End
         End
         Begin VB.Frame Help_Frame 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   860
            Left            =   133
            TabIndex        =   13
            Top             =   8280
            Width           =   1935
            Begin VB.Label lblWelComeScr 
               BackStyle       =   0  'Transparent
               Caption         =   "Welcome Screen"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   360
               MouseIcon       =   "frmMain.frx":40927
               MousePointer    =   99  'Custom
               TabIndex        =   21
               Top             =   75
               Width           =   1440
            End
            Begin VB.Label lblAboutETT 
               BackStyle       =   0  'Transparent
               Caption         =   "About ETT"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   360
               MouseIcon       =   "frmMain.frx":40C31
               MousePointer    =   99  'Custom
               TabIndex        =   20
               Top             =   480
               Width           =   1440
            End
            Begin VB.Image Image10 
               Height          =   240
               Left            =   60
               MouseIcon       =   "frmMain.frx":40F3B
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":41245
               Stretch         =   -1  'True
               Top             =   487
               Width           =   240
            End
            Begin VB.Image ImgWelComeScr 
               Height          =   240
               Left            =   60
               MouseIcon       =   "frmMain.frx":41E07
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":42111
               Stretch         =   -1  'True
               Top             =   82
               Width           =   240
            End
         End
         Begin VB.Frame Admin_Frame 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1460
            Left            =   133
            TabIndex        =   8
            Top             =   560
            Width           =   1935
            Begin VB.Label lblCompanyRef 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Company Ref"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   525
               MouseIcon       =   "frmMain.frx":43193
               MousePointer    =   99  'Custom
               TabIndex        =   12
               Top             =   1185
               Width           =   1170
            End
            Begin VB.Image Image4 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":4349D
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":437A7
               Stretch         =   -1  'True
               Top             =   1140
               Width           =   300
            End
            Begin VB.Label lblResetDB 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reset &DB Path"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   525
               MouseIcon       =   "frmMain.frx":440A6
               MousePointer    =   99  'Custom
               TabIndex        =   11
               Top             =   840
               Width           =   1230
            End
            Begin VB.Image Image2 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":443B0
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":446BA
               Stretch         =   -1  'True
               Top             =   780
               Width           =   300
            End
            Begin VB.Image Image6 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":449C4
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":44CCE
               Stretch         =   -1  'True
               Top             =   37
               Width           =   300
            End
            Begin VB.Image Image7 
               Height          =   300
               Left            =   60
               MouseIcon       =   "frmMain.frx":4866B
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":48975
               Stretch         =   -1  'True
               Top             =   435
               Width           =   300
            End
            Begin VB.Label lblLogDetail 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "L&og-In Detail"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   525
               MouseIcon       =   "frmMain.frx":49537
               MousePointer    =   99  'Custom
               TabIndex        =   10
               Top             =   465
               Width           =   1110
            End
            Begin VB.Label lblUserList 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&List Of Users"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   525
               MouseIcon       =   "frmMain.frx":49841
               MousePointer    =   99  'Custom
               TabIndex        =   9
               Top             =   90
               Width           =   1200
            End
         End
         Begin VB.Line Line9 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   1680
            X2              =   2160
            Y1              =   8040
            Y2              =   8040
         End
         Begin VB.Line Line8 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   1680
            X2              =   2160
            Y1              =   7695
            Y2              =   7695
         End
         Begin VB.Line Line7 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   1560
            X2              =   2160
            Y1              =   6720
            Y2              =   6720
         End
         Begin VB.Line Line6 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   1560
            X2              =   2160
            Y1              =   6360
            Y2              =   6360
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   1560
            X2              =   2160
            Y1              =   4605
            Y2              =   4605
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   1560
            X2              =   2160
            Y1              =   4920
            Y2              =   4905
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   1560
            X2              =   2160
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   2160
            X2              =   1200
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   120
            X2              =   0
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Image Nav_Image 
            Height          =   550
            Left            =   50
            MouseIcon       =   "frmMain.frx":49B4B
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":49E55
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   500
         End
         Begin VB.Label lblNav 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Navigation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Left            =   600
            MouseIcon       =   "frmMain.frx":4BB4F
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   2260
            Width           =   1065
         End
         Begin VB.Image Img_Nav 
            Height          =   255
            Left            =   30
            MouseIcon       =   "frmMain.frx":4BE59
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":4C163
            Stretch         =   -1  'True
            Top             =   2300
            Width           =   2145
         End
         Begin VB.Image Security_Image 
            Height          =   550
            Left            =   50
            MouseIcon       =   "frmMain.frx":4EDBF
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":4F0C9
            Stretch         =   -1  'True
            Top             =   6380
            Width           =   500
         End
         Begin VB.Label lblSecurity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Left            =   600
            MouseIcon       =   "frmMain.frx":4F3D3
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   6680
            Width           =   825
         End
         Begin VB.Image Img_Secur 
            Height          =   255
            Left            =   30
            MouseIcon       =   "frmMain.frx":4F6DD
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":4F9E7
            Stretch         =   -1  'True
            Top             =   6720
            Width           =   2145
         End
         Begin VB.Image Rpt_Image 
            Height          =   550
            Left            =   50
            MouseIcon       =   "frmMain.frx":52643
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":5294D
            Stretch         =   -1  'True
            Top             =   4620
            Width           =   500
         End
         Begin VB.Label lblReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Reports"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Left            =   600
            MouseIcon       =   "frmMain.frx":53217
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   4860
            Width           =   765
         End
         Begin VB.Image Img_Rpt 
            Height          =   255
            Left            =   30
            MouseIcon       =   "frmMain.frx":53521
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":5382B
            Stretch         =   -1  'True
            Top             =   4900
            Width           =   2145
         End
         Begin VB.Image Admin_Image 
            Height          =   550
            Left            =   50
            MouseIcon       =   "frmMain.frx":56487
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":56791
            Stretch         =   -1  'True
            Top             =   20
            Width           =   500
         End
         Begin VB.Image Help_Image 
            Height          =   550
            Left            =   50
            MouseIcon       =   "frmMain.frx":57B87
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":57E91
            Stretch         =   -1  'True
            Top             =   7740
            Width           =   500
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Help Info"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   600
            MouseIcon       =   "frmMain.frx":58B5B
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   8000
            Width           =   900
         End
         Begin VB.Image Img_Help 
            Height          =   255
            Left            =   30
            MouseIcon       =   "frmMain.frx":58E65
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":5916F
            Stretch         =   -1  'True
            Top             =   8040
            Width           =   2145
         End
         Begin VB.Label lblAdmin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Admin Only"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Left            =   600
            MouseIcon       =   "frmMain.frx":5BDCB
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   265
            Width           =   1140
         End
         Begin VB.Image Img_Admin 
            Height          =   255
            Left            =   30
            MouseIcon       =   "frmMain.frx":5C0D5
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":5C3DF
            Stretch         =   -1  'True
            Top             =   300
            Width           =   2145
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":5F069
         ScaleHeight     =   480
         ScaleWidth      =   2235
         TabIndex        =   3
         Top             =   0
         Width           =   2235
         Begin VB.Image Image12 
            Height          =   405
            Left            =   120
            Picture         =   "frmMain.frx":621CD
            Stretch         =   -1  'True
            Top             =   15
            Width           =   405
         End
      End
   End
   Begin VB.PictureBox COPYRIGHT_PIC 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   470
      Left            =   0
      Picture         =   "frmMain.frx":6266D
      ScaleHeight     =   465
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   10155
      Width           =   15240
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "M. SHAHID ASLAM MUGHAL.                  CONTACT # :  +92300-6410758"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   7440
         TabIndex        =   38
         Top             =   60
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Img_V1, Img_V2, lbl_V As Integer
Dim OptMenu As Boolean

Private Sub Initiate_Menu()
    Admin_Frame.Height = 0: Nav_Frame.Height = 0: Rpt_Frame.Height = 0: Secur_Frame.Height = 0: Help_Frame.Height = 0
    Admin_Frame.Visible = False: Nav_Frame.Visible = False: Rpt_Frame.Visible = False: Secur_Frame.Visible = False: Help_Frame.Visible = False
    
'    Img_V1 = 300: Img_V2 = 260: lbl_V = 50
    If Admin_Frame.Visible = False Then
        Nav_Image.Top = Img_Admin.Top + Img_V1: Img_Nav.Top = Nav_Image.Top + Img_V2
        lblNav.Top = Img_Nav.Top - lbl_V
    End If
    If Nav_Frame.Visible = False Then
        Rpt_Image.Top = Img_Nav.Top + Img_V1: Img_Rpt.Top = Rpt_Image.Top + Img_V2
        lblReport.Top = Img_Rpt.Top - lbl_V
    End If
    If Rpt_Frame.Visible = False Then
        Security_Image.Top = Img_Rpt.Top + Img_V1: Img_Secur.Top = Security_Image.Top + Img_V2
        lblSecurity.Top = Img_Secur.Top - lbl_V
    End If
    If Help_Frame.Visible = False Then
        Help_Image.Top = Img_Secur.Top + Img_V1: Img_Help.Top = Help_Image.Top + Img_V2
        lblHelp.Top = Img_Help.Top - lbl_V
    End If
End Sub

Private Sub Go_Admin()
    If Admin_Frame.Visible = False Then
        Nav_Image.Top = Img_Admin.Height + Img_V1
        lblNav.Top = Nav_Image.Top + 60
        Img_Nav.Top = lblNav.Top + Img_V2
    End If
    If Nav_Frame.Visible = False Then
        Rpt_Image.Top = Img_Nav.Top + Img_V1
        Img_Rpt.Top = Rpt_Image.Top + Img_V2
        lblReport.Top = Img_Rpt.Top - lbl_V
    End If
    If Rpt_Frame.Visible = False Then
        Security_Image.Top = Img_Rpt.Top + Img_V1: Img_Secur.Top = Security_Image.Top + Img_V2
        lblSecurity.Top = Img_Secur.Top - lbl_V
    End If
    If Help_Frame.Visible = False Then
        Help_Image.Top = Img_Secur.Top + Img_V1: Img_Help.Top = Help_Image.Top + Img_V2
        lblHelp.Top = Img_Help.Top - lbl_V
    End If
End Sub

Private Sub Go_Navigator()
    If Rpt_Frame.Visible = False Then
        Security_Image.Top = Img_Rpt.Top + Img_V1: Img_Secur.Top = Security_Image.Top + Img_V2
        lblSecurity.Top = Img_Secur.Top - lbl_V
    End If
    If Help_Frame.Visible = False Then
        Help_Image.Top = Img_Secur.Top + Img_V1: Img_Help.Top = Help_Image.Top + Img_V2
        lblHelp.Top = Img_Help.Top - lbl_V
    End If
End Sub

Private Sub Go_Reports()
    If Help_Frame.Visible = False Then
        Help_Image.Top = Img_Secur.Top + Img_V1: Img_Help.Top = Help_Image.Top + Img_V2
        lblHelp.Top = Img_Help.Top - lbl_V
    End If
End Sub

Private Sub Admin_Timer1_Timer()
    If Admin_Frame.Height <= 1460 Then
        Admin_Frame.Height = Admin_Frame.Height + 20
    ElseIf Admin_Frame.Height >= 1460 Then
        Admin_Timer1.Interval = 0
    End If
End Sub

Private Sub Admin_Timer2_Timer()
    If Admin_Frame.Height <> 15 Then
        Admin_Frame.Height = Admin_Frame.Height - 20
    ElseIf Admin_Frame.Height <= 15 Then
        Admin_Timer2.Interval = 0: Admin_Frame.Visible = False
    End If
End Sub

Private Sub Help_Timer1_Timer()
    If Help_Frame.Height <= 860 Then
        Help_Frame.Height = Help_Frame.Height + 20
    ElseIf Help_Frame.Height >= 860 Then
        Help_Timer1.Interval = 0
    End If
End Sub

Private Sub Help_Timer2_Timer()
    If Help_Frame.Height <> 15 Then
        Help_Frame.Height = Help_Frame.Height - 20
    ElseIf Help_Frame.Height <= 15 Then
        Help_Timer2.Interval = 0: Help_Frame.Visible = False
    End If
End Sub

Private Sub Img_Admin_Click()
    Call Go_Admin 'Call Admin Menu Sidebar.
    If Admin_Frame.Height <= 15 Then Admin_Frame.Visible = True: Admin_Timer1.Interval = 2: Exit Sub
    If Admin_Frame.Height >= 1460 Then Admin_Timer2.Interval = 2: Exit Sub
End Sub

Private Sub Img_Help_Click()
    If Help_Frame.Height <= 15 Then Help_Frame.Visible = True: Help_Timer1.Interval = 2: Exit Sub
    If Help_Frame.Height >= 860 Then Help_Timer2.Interval = 2: Exit Sub
End Sub

Private Sub Img_Nav_Click()
    If Nav_Frame.Height <= 15 Then Nav_Frame.Visible = True: Nav_Timer1.Interval = 2: Exit Sub
    If Nav_Frame.Height >= 1460 Then Nav_Timer2.Interval = 2: Exit Sub
End Sub

Private Sub Img_Rpt_Click()
    If Rpt_Frame.Height <= 15 Then Rpt_Frame.Visible = True: Rpt_Timer1.Interval = 2: Exit Sub
    If Rpt_Frame.Height >= 1460 Then Rpt_Timer2.Interval = 2: Exit Sub
End Sub

Private Sub Img_Secur_Click()
    If Secur_Frame.Height <= 15 Then Secur_Frame.Visible = True: Secur_Timer1.Interval = 2: Exit Sub
    If Secur_Frame.Height >= 740 Then Secur_Timer2.Interval = 2: Exit Sub
End Sub

Private Sub ImgWelComeScr_Click()
    Call lblWelComeScr_Click
End Sub

Private Sub lblCompanyRef_Click()
    Call Ctrl_ET.msg_Consutruct 'Under Construction.
End Sub

Private Sub lblDaily_Month_Rpt_Click()
    Call Ctrl_ET.msg_Consutruct 'Under Construction.
End Sub

Private Sub lblDoneJob_Rpt_Click()
    Call Ctrl_ET.msg_Consutruct 'Under Construction.
End Sub

Private Sub lblJobpending_Rpt_Click()
    Call Ctrl_ET.msg_Consutruct 'Under Construction.
End Sub

Private Sub lblCheckStatus_Click()
    Load frmCheckStatus: frmCheckStatus.Show
End Sub

Private Sub lblAboutETT_Click()
    Load frmAbout: frmAbout.Show
End Sub

Private Sub lblAirLine_Click()
    Load frmAirLine: frmAirLine.Show
End Sub

Private Sub lblChangePwd_Click()
    Load frmChangePwd: frmChangePwd.Show
End Sub

Private Sub lblCharges_Click()
    Load frmCharges: frmCharges.Show
End Sub

Private Sub lblDestination_Click()
    Load frmDestination: frmDestination.Show
End Sub

Private Sub lblLogDetail_Click()
    Load frmLogRecord: frmLogRecord.Show
End Sub

Private Sub lblNewJob_Click()
    Load frmCustomer: frmCustomer.Show
End Sub

Private Sub lblNewUser_Click()
    Load frmUser_Create: frmUser_Create.Show
End Sub

Private Sub lblResetDB_Click()
    Call Ctrl_ET.msg_Consutruct 'Under Construction.
End Sub

Private Sub lblServices_Click()
    Load frmServices: frmServices.Show
End Sub

Private Sub lblUserList_Click()
    Load frmUserList: frmUserList.Show
End Sub

Private Sub lblWelComeScr_Click()
    Load frmSplash: frmSplash.Show
End Sub

Private Sub MDIForm_Load()
'    Img_V1 = 300: Img_V2 = 260: lbl_V = 50
'    Call Initiate_Menu
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim Rst As New ADODB.Recordset 'For Login Information in Login Table.
    Rst.Open "SELECT * FROM tblLogIn_Record", DB_Conect, adOpenStatic, adLockOptimistic
    With Rst
        .Close: .Open "SELECT * FROM tblLogIn_Record WHERE User_ID='" & lblUserID & "' AND Login_Date='" & lblDate & "'"
            .Fields(3).Value = Date: .Fields(4).Value = Time
            .Update
    End With
    Rst.Close
End Sub

Private Sub Nav_Timer1_Timer()
    If Nav_Frame.Height <= 1695 Then
        Nav_Frame.Height = Nav_Frame.Height + 20
    ElseIf Nav_Frame.Height >= 1695 Then
        Nav_Timer1.Interval = 0
    End If
End Sub

Private Sub Nav_Timer2_Timer()
    If Nav_Frame.Height <> 15 Then
        Nav_Frame.Height = Nav_Frame.Height - 20
    ElseIf Nav_Frame.Height <= 15 Then
        Nav_Timer2.Interval = 0: Nav_Frame.Visible = False
    End If
End Sub

Private Sub Rpt_Timer1_Timer()
    If Rpt_Frame.Height <= 1575 Then
        Rpt_Frame.Height = Rpt_Frame.Height + 20
    ElseIf Rpt_Frame.Height >= 1575 Then
        Rpt_Timer1.Interval = 0
    End If
End Sub

Private Sub Rpt_Timer2_Timer()
    If Rpt_Frame.Height <> 15 Then
        Rpt_Frame.Height = Rpt_Frame.Height - 20
    ElseIf Rpt_Frame.Height <= 15 Then
        Rpt_Timer2.Interval = 0: Rpt_Frame.Visible = False
    End If
End Sub

Private Sub Secur_Timer1_Timer()
    If Secur_Frame.Height <= 740 Then
        Secur_Frame.Height = Secur_Frame.Height + 20
    ElseIf Secur_Frame.Height >= 740 Then
        Secur_Timer1.Interval = 0
    End If
End Sub

Private Sub Secur_Timer2_Timer()
    If Secur_Frame.Height <> 15 Then
        Secur_Frame.Height = Secur_Frame.Height - 20
    ElseIf Secur_Frame.Height <= 15 Then
        Secur_Timer2.Interval = 0: Secur_Frame.Visible = False
    End If
End Sub
