VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmCustomer.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox txtNIC 
      Height          =   330
      Left            =   3360
      TabIndex        =   45
      Top             =   1260
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####-#######-#"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtBalance 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7680
      TabIndex        =   42
      Text            =   "txtBalance"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtAdvance 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      TabIndex        =   41
      Text            =   "txtAdvance"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtGrandTot 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   38
      Text            =   "txtGrandTot"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtServCharges 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   36
      Text            =   "txtServCharges"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtAirCharges 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   35
      Text            =   "txtAirCharges"
      Top             =   4800
      Width           =   1335
   End
   Begin ctrlNSDataCombo.NSDataCombo NSDDestination 
      Height          =   255
      Left            =   6480
      TabIndex        =   33
      Top             =   4280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   6195
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   16711680
      FCOLO           =   16711935
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCustomer.frx":C668
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   3
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdCancel 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   6195
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   16711680
      FCOLO           =   16711935
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCustomer.frx":C684
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   3
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdSubmit 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   6195
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Submit"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   16711680
      FCOLO           =   16711935
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCustomer.frx":C6A0
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   3
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   6200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&New Record"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12648384
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCustomer.frx":C6BC
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin MSComCtl2.DTPicker Date_Expected 
      Height          =   315
      Left            =   6480
      TabIndex        =   26
      Top             =   4680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   56033281
      CurrentDate     =   38959
   End
   Begin ctrlNSDataCombo.NSDataCombo NSDCAirLine 
      Height          =   255
      Left            =   2040
      TabIndex        =   25
      Top             =   4280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ctrlNSDataCombo.NSDataCombo NSDOrigin 
      Height          =   255
      Left            =   6480
      TabIndex        =   24
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ctrlNSDataCombo.NSDataCombo NSDCService 
      Height          =   255
      Left            =   2040
      TabIndex        =   23
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtMobNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      TabIndex        =   18
      Text            =   "txtMobNo"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtHomeTel 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   17
      Text            =   "txtHomeTel"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtPostAdd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "frmCustomer.frx":C6D8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtFName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   15
      Text            =   "txtFName"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   14
      Text            =   "txtName"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtPassportNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   8
      Text            =   "txtPassportNo"
      Top             =   900
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   315
      Left            =   7080
      TabIndex        =   7
      Top             =   1260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   56033281
      CurrentDate     =   38959
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M Shahid Aslam Mughal. +923006410758"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   6600
      TabIndex        =   46
      Top             =   5160
      Width           =   2400
   End
   Begin VB.Label lblApplicantNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblApplicantNo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   7560
      TabIndex        =   44
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant No : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6000
      TabIndex        =   43
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Balance : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6520
      TabIndex        =   40
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Advance : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3640
      TabIndex        =   39
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   37
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lblDestination 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDestination"
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
      Height          =   255
      Left            =   8400
      TabIndex        =   34
      Top             =   4280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destination : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   4280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Services Charges : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -120
      TabIndex        =   31
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air Line Charges : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   30
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblServices 
      BackStyle       =   0  'Transparent
      Caption         =   "lblServices"
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
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAirLine 
      BackStyle       =   0  'Transparent
      Caption         =   "lblAirLine"
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
      Height          =   255
      Left            =   3960
      TabIndex        =   28
      Top             =   4280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblOrigin 
      BackStyle       =   0  'Transparent
      Caption         =   "lblOrigin"
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
      Height          =   255
      Left            =   8400
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air Line Name  : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4280
      Width           =   1695
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Expected Date : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Origin : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Services Name : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile # : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Home Tel # : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Address : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NIC # : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Passport # : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   900
      Width           =   1455
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstSrv As New ADODB.Recordset 'For Services DB-Table.
Dim RstArL As New ADODB.Recordset 'For Air Line DB-Table.
Dim RstCont As New ADODB.Recordset 'For Contry DB-Table.
Dim RstChar As New ADODB.Recordset 'For Air Line Charges.
Dim RstCust_Info As New ADODB.Recordset 'For Customer Information DB-Table.
Dim RstCust_Pay As New ADODB.Recordset 'For Customer Payment DB-Table.

Private Sub Populate_Save()
    With RstCust_Info
        .AddNew
            .Fields(0).Value = lblApplicantNo
            .Fields(1).Value = txtPassportNo.Text: .Fields(2).Value = txtNIC.Text
            .Fields(3).Value = txtName.Text: .Fields(4).Value = txtFName.Text
            .Fields(5).Value = txtPostAdd.Text: .Fields(6).Value = txtHomeTel.Text
            .Fields(7).Value = txtMobNo.Text: .Fields(8).Value = DTPicker.Value
            .Fields(9).Value = lblServices: .Fields(10).Value = lblDestination
            .Fields(11).Value = lblAirLine: .Fields(12).Value = 0
        .Update
        
        With RstCust_Pay
            .AddNew
            .Fields(0).Value = lblApplicantNo
            .Fields(1).Value = Val(txtAirCharges.Text): .Fields(2).Value = Val(txtServCharges.Text)
            .Fields(3).Value = Val(txtGrandTot.Text): .Fields(4).Value = Val(txtAdvance.Text)
            .Fields(5).Value = Val(txtBalance.Text)
            If txtBalance.Text <= 0 Then .Fields(6).Value = 1 'Okay(Ready)
            If txtBalance.Text > 0 Then .Fields(6).Value = 0 'Not Okay(Not Ready)
            .Update
        End With: MsgBox "Record has been Saved Successfully", vbInformation, "Saving Record"
    End With
End Sub

Private Sub InitNSD()
    'For Services List In NS-Data Combo box.
    With NSDCService
        .ClearColumn: .AddColumn "Code", 1200: .AddColumn "Service", 2200
        
        .Connection = DB_Conect.ConnectionString
        .sqlFields = "Srv_ID, Srv_Des": .sqlTables = "tblServices"
        .sqlSortOrder = "Srv_ID ASC": .BoundField = "Srv_ID"
        
        .PageBy = 25
        .DisplayCol = 2 'That Value Want to Display In Field(TextBox).
        
        RstSQL = "SELECT * FROM tblServices WHERE Srv_Des='" & NSDCService.Text & "'"
        RstSrv.Close: RstSrv.Open RstSQL
        If RstSrv.RecordCount > 0 Then lblServices = RstSrv.Fields(0).Value
            
'            If ((NSDCService.Text <> "") And (NSDCAirLine.Text <> "") And (NSDDestination.Text <> "") And (NSDOrigin.Text <> "")) Then
'                RstChar.Close: RstChar.Open "SELECT * FROM tblCharges WHERE AL_ID='" & lblAirLine & "' AND Srv_ID='" & lblServices & "' AND Origin='" & NSDOrigin.Text & "' AND Destination='" & NSDDestination.Text & "'"
'                If RstChar.RecordCount > 0 Then txtAirCharges.Text = RstChar.Fields(4).Value: txtServCharges.Text = RstChar.Fields(5).Value
'            End If
        If RstSrv.RecordCount <= 0 Then lblServices = ""
        
        .setDropWindowSize 3600, 3000
        .TextReadOnly = True: .SetDropDownTitle = "Services List"
    End With
'==========================================================================================
    'For Air Line List In NS-Data Combo box.
    With NSDCAirLine
        .ClearColumn: .AddColumn "Code", 1200: .AddColumn "Air Line Name", 2200
        
        .Connection = DB_Conect.ConnectionString
        .sqlFields = "AL_ID, AL_Des": .sqlTables = "tblAirLine"
        .sqlSortOrder = "AL_ID ASC": .BoundField = "AL_ID"
        
        .PageBy = 25
        .DisplayCol = 2 'That Value Want to Display In Field(TextBox).
        
        RstSQL = "SELECT * FROM tblAirLine WHERE AL_Des='" & NSDCAirLine.Text & "'"
        RstArL.Close: RstArL.Open RstSQL
        If RstArL.RecordCount > 0 Then
            lblAirLine = RstArL.Fields(0).Value
            
            If ((NSDCService.Text <> "") And (NSDCAirLine.Text <> "") And (NSDDestination.Text <> "") And (NSDOrigin.Text <> "")) Then
                RstChar.Close: RstChar.Open "SELECT * FROM tblCharges WHERE AL_ID='" & lblAirLine & "' AND Srv_ID='" & lblServices & "' AND Origin='" & NSDOrigin.Text & "' AND Destination='" & NSDDestination.Text & "'"
                If RstChar.RecordCount > 0 Then
                    txtAirCharges.Text = RstChar.Fields(4).Value
                    txtServCharges.Text = RstChar.Fields(5).Value
                ElseIf RstChar.RecordCount <= 0 Then
                    txtAirCharges.Text = "0"
                    txtServCharges.Text = "0"
                End If
            End If
        End If
        
        If RstArL.RecordCount <= 0 Then lblAirLine = ""
        
        .setDropWindowSize 3600, 3000
        .TextReadOnly = True: .SetDropDownTitle = "Air Line List"
    End With
'==========================================================================================
    'For Origin Country List In NS-Data Combo box.
    With NSDOrigin
        .ClearColumn: .AddColumn "Country Code", 1200: .AddColumn "Country Name", 2200
        
        .Connection = DB_Conect.ConnectionString
        .sqlFields = "Country_ID, Country_Name": .sqlTables = "tblDestination"
        .sqlSortOrder = "Country_Name ASC": .BoundField = "Country_ID"
        
        .PageBy = 25
        .DisplayCol = 2 'That Value Want to Display In Field(TextBox).
        
        RstSQL = "SELECT * FROM tblDestination WHERE Country_Name='" & NSDOrigin.Text & "'"
        RstCont.Close: RstCont.Open RstSQL
        If RstCont.RecordCount > 0 Then
            lblOrigin = RstCont.Fields(0).Value

            If ((NSDCService.Text <> "") And (NSDCAirLine.Text <> "") And (NSDDestination.Text <> "") And (NSDOrigin.Text <> "")) Then
                RstChar.Close: RstChar.Open "SELECT * FROM tblCharges WHERE AL_ID='" & lblAirLine & "' AND Srv_ID='" & lblServices & "' AND Origin='" & NSDOrigin.Text & "' AND Destination='" & NSDDestination.Text & "'"
                If RstChar.RecordCount > 0 Then
                    txtAirCharges.Text = RstChar.Fields(4).Value
                    txtServCharges.Text = RstChar.Fields(5).Value
                ElseIf RstChar.RecordCount <= 0 Then
                    txtAirCharges.Text = "0"
                    txtServCharges.Text = "0"
                End If
            End If
        End If
        
        If RstCont.RecordCount <= 0 Then lblOrigin = ""
        
        .setDropWindowSize 3600, 3000
        .TextReadOnly = True: .SetDropDownTitle = "Country List"
    End With
'==========================================================================================
    'For Destination Country List In NS-Data Combo box.
    With NSDDestination
        .ClearColumn: .AddColumn "Country Code", 1200: .AddColumn "Country Name", 2200
        
        .Connection = DB_Conect.ConnectionString
        .sqlFields = "Country_ID, Country_Name": .sqlTables = "tblDestination"
        .sqlSortOrder = "Country_Name ASC": .BoundField = "Country_ID"
        
        .PageBy = 25
        .DisplayCol = 2 'That Value Want to Display In Field(TextBox).
        
        RstSQL = "SELECT * FROM tblDestination WHERE Country_Name='" & NSDDestination.Text & "'"
        RstCont.Close: RstCont.Open RstSQL
        If RstCont.RecordCount > 0 Then
            lblDestination = RstCont.Fields(0).Value
        
            If ((NSDCService.Text <> "") And (NSDCAirLine.Text <> "") And (NSDDestination.Text <> "") And (NSDOrigin.Text <> "")) Then
                RstChar.Close: RstChar.Open "SELECT * FROM tblCharges WHERE AL_ID='" & lblAirLine & "' AND Srv_ID='" & lblServices & "' AND Origin='" & NSDOrigin.Text & "' AND Destination='" & NSDDestination.Text & "'"
                If RstChar.RecordCount > 0 Then
                    txtAirCharges.Text = RstChar.Fields(4).Value
                    txtServCharges.Text = RstChar.Fields(5).Value
                ElseIf RstChar.RecordCount <= 0 Then
                    txtAirCharges.Text = "0"
                    txtServCharges.Text = "0"
                End If
            End If
        End If
        
        If RstCont.RecordCount <= 0 Then lblDestination = ""
        
        .setDropWindowSize 3600, 3000
        .TextReadOnly = True: .SetDropDownTitle = "Country List"
    End With
'==========================================================================================
End Sub


Private Sub CmdCancel_Click()
    Call Ctrl_ET.Populate_Entery(frmCustomer, False) 'Restrict User for Enteries in Form.
    CmdNew.Enabled = True: CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    Call Ctrl_ET.Populate_Text_Clear(frmCustomer) 'To Clearing The Text.
    txtNIC.Mask = ""
    txtAdvance.Text = "0": CmdNew.SetFocus 'Set the Focus to related Button.
End Sub

Private Sub CmdExit_Click()
    Unload frmCustomer
End Sub

Private Sub CmdNew_Click()
    Call Ctrl_ET.Populate_Entery(frmCustomer, True) 'Allow User for Enteries in Form.
    CmdNew.Enabled = False: CmdSubmit.Enabled = True: CmdCancel.Enabled = True
    Call Ctrl_ET.Populate_Text_Clear(frmCustomer) 'To Clearing The Text.
    
    txtPassportNo.SetFocus 'Set the Focus to related TextBox.
End Sub

Private Sub CmdSubmit_Click()
    If txtPassportNo.Text = "" Then _
        MsgBox "Please enter the valid Passport Number.", vbCritical, "Passport Error!": _
               SendKeys "{Home}+{End}": txtPassportNo.SetFocus: Exit Sub
    
    If txtNIC.Text = "" Then _
        MsgBox "Please enter the Vlid NIC Number", vbCritical, "NIC Number Error!": _
               SendKeys "{Home}+{End}": txtNIC.SetFocus: Exit Sub
               
    If txtServCharges.Text = "" Then _
        MsgBox "Invalid Services Charges Please Contact Programe Vendor", vbCritical, "Service Charges Error!": _
                Call CmdCancel_Click: Exit Sub

    If txtAirCharges.Text = "" Then _
        MsgBox "Invalid Air Line Charges Please Contact Programe Vendor", vbCritical, "Air Line Charges Error!": _
                Call CmdCancel_Click: Exit Sub
    
    If NSDCService.Text = "" Then _
        MsgBox "Please select the valid Services Provide.", vbCritical, "Services Error!": _
                NSDCService.SetFocus: Exit Sub
               
    If NSDCAirLine.Text = "" Then _
        MsgBox "Please select valid Air Line.", vbCritical, "Air Line Error!": _
               SendKeys "{Home}+{End}": NSDCAirLine.SetFocus: Exit Sub
    
    If NSDOrigin.Text = "" Then _
        MsgBox "Please select valid Origin.", vbCritical, "Origin Error!": _
               NSDOrigin.SetFocus: Exit Sub
               
    If NSDDestination.Text = "" Then _
        MsgBox "Please select valid Destination.", vbCritical, "Destination Error!": _
               NSDDestination.SetFocus: Exit Sub
    
    If NSDOrigin.Text = NSDDestination.Text Then _
        MsgBox "Origin and Destination must be different." & vbCrLf & _
               "Please try again.", vbCritical, "Origin/Destination Error!": _
               SendKeys "{Home}+{End}": NSDDestination.SetFocus: Exit Sub
    
    Call Populate_Save 'Call for Saving Records in Database.
    Call Ctrl_ET.Populate_Text_Clear(frmCustomer) 'Clearing the TextBoxes.
    Call Ctrl_ET.Populate_Entery(frmCustomer, False) 'Restrict the Enteries.
    CmdNew.Enabled = True: CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    CmdNew.SetFocus
End Sub

Private Sub Date_Expected_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{Home}+{End}": txtAdvance.SetFocus
    End If
End Sub

Private Sub DTPicker_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtName.SetFocus
        Date_Expected.Value = Date + 10
    End If
End Sub

Private Sub Form_Activate()
    Call InitNSD 'For Initializing Data Combo.
End Sub

Private Sub Form_Load()
    frmCustomer.Move (frmMain.Width / 6), (frmMain.Height / 12)
    Call Ctrl_ET.Populate_Text_Clear(frmCustomer) 'To Clearing The Text.
    txtAdvance.Text = "0"
    
    RstSQL = "SELECT * FROM tblServices" 'For Services that Travelling Agency Provides.
    RstSrv.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM tblAirLine" 'For Air Lines With Which Travelling Agency Dealing.
    RstArL.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM tblDestination" 'For Destination Countries.
    RstCont.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM tblCustomer_Info" 'For Customer Information.
    RstCust_Info.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM tblCustomer_Payment" 'For Customer Detail.
    RstCust_Pay.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM tblCharges" 'For Air Line Charges Detail.
    RstChar.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    DTPicker = Date: Call InitNSD 'For Initializing Data Combo.
    Call Ctrl_ET.Populate_Entery(frmCustomer, False) 'Restrict User for Enteries in Form.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstSrv.Close: RstArL.Close: RstCont.Close 'Closing DB-Table.
    RstCust_Info.Close: RstCust_Pay.Close 'Closing DB-Table.
    RstChar.Close 'Closing DB-Table.
End Sub

Private Sub NSDCAirLine_Change()
    Call InitNSD 'Call For NSDCServices Text Box.
End Sub

Private Sub NSDDestination_Change()
    Call InitNSD 'Call For NSDCServices Text Box.
End Sub

Private Sub NSDOrigin_Change()
    Call InitNSD 'Call For NSDCServices Text Box.
End Sub

Private Sub NSDCService_Change()
    Call InitNSD 'Call For NSDCServices Text Box.
End Sub

Private Sub txtAdvance_Change()
    txtBalance.Text = Val(txtGrandTot.Text) - Val(txtAdvance.Text)
End Sub

Private Sub txtAdvance_KeyPress(KeyAscii As Integer)
    Call Ctrl_ET.Populate_NumOnly(KeyAscii) 'Call for Numeric Values Only.
    If KeyAscii = 13 Then
        SendKeys "{Home}+{End}": txtBalance.SetFocus
    End If
End Sub

Private Sub txtAirCharges_Change()
    txtGrandTot.Text = Val(txtAirCharges.Text) + Val(txtServCharges.Text)
    txtAdvance.Text = "0"
End Sub

Private Sub txtAirCharges_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Home}+{End}": txtServCharges.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtBalance_Change()
    txtBalance = Val(txtGrandTot.Text) - Val(txtAdvance.Text)
End Sub

Private Sub txtBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Home}+{End}": CmdSubmit.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtFName_KeyPress(KeyAscii As Integer)
    Call Ctrl_ET.Populate_Alpha_Char(KeyAscii, frmCustomer, txtFName)
    Call Ctrl_ET.Populate_CharOnly(KeyAscii) 'Character Values only.
    If KeyAscii = 13 Then
        If txtFName.Text <> "" Then txtPostAdd.SetFocus
    End If
End Sub

Private Sub txtGrandTot_Change()
    txtBalance.Text = Val(txtGrandTot.Text) - Val(txtAdvance.Text)
End Sub

Private Sub txtGrandTot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Home}+{End}": txtAdvance.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtHomeTel_KeyPress(KeyAscii As Integer)
    Call Ctrl_ET.Populate_NumOnly(KeyAscii) 'Numeric Values only.
    If KeyAscii = 13 Then
        If txtHomeTel.Text <> "" Then txtMobNo.SetFocus
    End If
End Sub

Private Sub txtMobNo_KeyPress(KeyAscii As Integer)
    Call Ctrl_ET.Populate_NumOnly(KeyAscii) 'Numeric Values only.
    If KeyAscii = 13 Then
        If txtMobNo.Text <> "" Then NSDCService.SetFocus
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Call Ctrl_ET.Populate_Alpha_Char(KeyAscii, frmCustomer, txtName)
    Call Ctrl_ET.Populate_CharOnly(KeyAscii) 'Character Values only.
    If KeyAscii = 13 Then
        If txtName.Text <> "" Then txtFName.SetFocus
    End If
End Sub

Private Sub txtNIC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNIC.Text <> "" Then DTPicker.SetFocus
    End If
End Sub

Private Sub txtPassportNo_KeyPress(KeyAscii As Integer)
    Call Ctrl_ET.Populate_AutoID(RstCust_Info) 'Call for Auto Applicant Number.
    lblApplicantNo = Auto_ID 'Display Auto Applicant Number.
    If KeyAscii = 13 Then
        If txtPassportNo.Text <> "" Then txtNIC.SetFocus
    End If
End Sub

Private Sub txtPostAdd_KeyPress(KeyAscii As Integer)
    Call Ctrl_ET.Populate_Alpha_Char(KeyAscii, frmCustomer, txtPostAdd)
    If KeyAscii = 13 Then
        If txtPostAdd.Text <> "" Then KeyAscii = 0: txtHomeTel.SetFocus
    End If
End Sub

Private Sub txtServCharges_Change()
    txtGrandTot.Text = Val(txtAirCharges.Text) + Val(txtServCharges.Text)
End Sub

Private Sub txtServCharges_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Home}+{End}": txtGrandTot.SetFocus
    Else
        KeyAscii = 0
    End If
End Sub
