VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About System ..."
   ClientHeight    =   5025
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   5790
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   3468.344
   ScaleMode       =   0  'User
   ScaleWidth      =   5437.11
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000006&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   30
         Picture         =   "frmAbout.frx":8462
         Stretch         =   -1  'True
         Top             =   30
         Width           =   960
      End
   End
   Begin VB.Label lblvoting 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't forget about voting, im waiting your comments please leave your comments whatever . . . . . . . . . . ."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1065
      Left            =   120
      TabIndex        =   10
      Top             =   3885
      Width           =   2175
   End
   Begin VB.Label lblmobile 
      BackStyle       =   0  'Transparent
      Caption         =   "+92-300-6410758."
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   1440
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "methoomirza@hotmail.com"
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   1440
      TabIndex        =   8
      Top             =   3360
      Width           =   2190
   End
   Begin VB.Label lblcontactinfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Information : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblDesignation 
      BackStyle       =   0  'Transparent
      Caption         =   "Working in the I.T Department                   as Data Processor Additional Programmer."
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label lblDescribtion 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":6B444
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   690
      TabIndex        =   4
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label lblMyName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M Shahid Aslam Mughal."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3120
      TabIndex        =   3
      Top             =   2160
      Width           =   2235
   End
   Begin VB.Label lblProgrammedby 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designing and Programmed by : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2850
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   1140
   End
   Begin VB.Label lblFirstLine 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Created in mostly using Windows Development kit "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   5025
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormslblValues()
    lblFirstLine = " Created in mostly using Windows Development kit "
    lblDescribtion = "This software was design for Air Tickting System purpose only. It will record / manage your customer and do some reporting on your demand."
    lblProgrammedby = "Designing and Programmed by : ": lblMyName = "M Shahid Aslam Mughal."
    lblDesignation = "Working in the I.T Department                   as Data Processor Additional Programmer."
    lblcontactinfo = "Contact Information : ": lblmail = "methoomirza@hotmail.com": lblmobile = "+92-300-6410758."
    lblvoting.Caption = "Don't forget about voting, im waiting your comments please leave your comments whatever . . . . . . . . . . ."
End Sub

Private Sub Form_Click()
   Unload frmAbout
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload frmAbout
    End If
End Sub

Private Sub Form_Load()
    frmAbout.Move (frmMain.Width / 3), (frmMain.Height / 6): Call FormslblValues 'For Label's Entries.
End Sub
