VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCheckStatus 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M Shahid Aslam Mughal. +923006410758"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
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
   Picture         =   "frmCheckStatus.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNIC 
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
      Left            =   2520
      TabIndex        =   27
      Text            =   "txtNIC"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
      Begin VB.Label lblReturnDate 
         BackStyle       =   0  'Transparent
         Caption         =   "lblReturnDate"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Return Date : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblMobile 
         BackStyle       =   0  'Transparent
         Caption         =   "lblMobile"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile # : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblStatus"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   75
         TabIndex        =   22
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   705
         Width           =   1695
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "lblName"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   705
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Father Name : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label lblFName 
         BackStyle       =   0  'Transparent
         Caption         =   "lblFName"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   1005
         Width           =   2535
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Land Line # : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   1395
         Width           =   1695
      End
      Begin VB.Label lblTelNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "lblTelNumber"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1395
         Width           =   2535
      End
      Begin VB.Label lblSubmitDate 
         BackStyle       =   0  'Transparent
         Caption         =   "lblSubmitDate"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   40
         Width           =   2535
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Submit Date : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   40
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount : "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Advance : "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance : "
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblAmount"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   3300
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblAdvance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblAdvance"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   3300
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblBalance"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   3300
         TabIndex        =   7
         Top             =   2880
         Width           =   1215
      End
   End
   Begin LVbuttons.LaVolpeButton CmdReady 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Ready"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCheckStatus.frx":4EC7
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
   Begin VB.TextBox txtPassport 
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
      Left            =   2520
      TabIndex        =   3
      Text            =   "txtPassport"
      Top             =   840
      Width           =   2175
   End
   Begin LVbuttons.LaVolpeButton CmdOk 
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Ok"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCheckStatus.frx":4EE3
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
   Begin VB.Label Label14 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   450
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label3 
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
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Status "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1433
      TabIndex        =   0
      Top             =   0
      Width           =   2145
   End
End
Attribute VB_Name = "frmCheckStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstPenRpt As New ADODB.Recordset
Dim RstPayment As New ADODB.Recordset

Public Sub Populate_Rec()
    With RstPenRpt
        txtPassport.Text = .Fields(1).Value: txtNIC.Text = .Fields(2).Value
        lblSubmitDate = .Fields(8).Value: lblReturnDate = .Fields(9).Value
        lblName = .Fields(3).Value: lblFName = .Fields(4).Value
        lblTelNumber = .Fields(6).Value: lblMobile = .Fields(7).Value
        
        With RstPayment
            lblAmount = .Fields(3).Value: lblAdvance = .Fields(4).Value
            lblBalance = .Fields(5).Value
        End With
        CmdOk.Enabled = True
    End With
End Sub

Private Sub CmdOk_Click()
    Unload frmCheckStatus 'Unload the Current Form.
End Sub

Private Sub Form_Load()
    frmCheckStatus.Move (frmMain.Width / 3), (frmMain.Height / 10)
    Call Ctrl_ET.Populate_Text_Clear(frmCheckStatus) 'To Initialize Text Boxes.
    lblSubmitDate = "---": lblName = "---": lblFName = "---": lblTelNumber = "---"
    lblAmount = "0": lblAdvance = "0": lblBalance = "0": lblStatus = ""
    
    RstPenRpt.Open "SELECT * FROM tblCustomer_Info", DB_Conect, adOpenStatic, adLockOptimistic
    RstPayment.Open "SELECT * FROM tblCustomer_Payment", DB_Conect, adOpenStatic, adLockOptimistic
    
    RstPenRpt.Close: RstPenRpt.Open "SELECT * FROM tblCustomer_Info WHERE PassPort_No='" & Passport & "'"
    RstPayment.Close: RstPayment.Open "SELECT * FROM tblCustomer_Payment WHERE Agency_ID=" & Catch_Field
    Call Populate_Rec 'Call for Record Display in Fields.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstPenRpt.Close: RstPayment.Close
End Sub
