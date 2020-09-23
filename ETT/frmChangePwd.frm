VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmChangePwd 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Users Change Password . . . . "
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Picture         =   "frmChangePwd.frx":0000
   ScaleHeight     =   4170
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   12582912
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmChangePwd.frx":519D
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
   Begin LVbuttons.LaVolpeButton CmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   12582912
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmChangePwd.frx":51B9
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
   Begin LVbuttons.LaVolpeButton CmdChange 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "C&hange"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   12582912
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmChangePwd.frx":51D5
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
   Begin VB.TextBox txtConPwd 
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
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "txtConPwd"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtNewPwd 
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
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "txtNewPwd"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtPwd 
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
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "txtPwd"
      Top             =   1900
      Width           =   2535
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
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   2040
      TabIndex        =   14
      Top             =   3480
      Width           =   4080
   End
   Begin VB.Label lblUserID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblUserID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   330
      Left            =   3840
      TabIndex        =   11
      Top             =   1500
      Width           =   2535
   End
   Begin VB.Label lblUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblUserName"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   330
      Left            =   3840
      TabIndex        =   10
      Top             =   1040
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm New Password : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1215
      TabIndex        =   9
      Top             =   2880
      Width           =   2520
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "New Password : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1215
      TabIndex        =   8
      Top             =   2400
      Width           =   2520
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1215
      TabIndex        =   7
      Top             =   1900
      Width           =   2520
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User ID : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1215
      TabIndex        =   6
      Top             =   1500
      Width           =   2520
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1215
      TabIndex        =   5
      Top             =   1040
      Width           =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Change Password "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   2280
      TabIndex        =   4
      Top             =   60
      Width           =   3420
   End
End
Attribute VB_Name = "frmChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstLog As New ADODB.Recordset

Private Sub CmdCancel_Click()
    Call Ctrl_ET.Populate_Text_Clear(frmChangePwd) 'Call for Clearing the Text Boxes.
    Unload frmChangePwd
End Sub

Private Sub CmdChange_Click()
    If txtPwd.Text <> "" Then
        RstLog.Close: RstLog.Open "SELECT * FROM tblLogin WHERE User_ID='" & frmMain.lblUserID & "'"
        
        If txtPwd.Text = RstLog.Fields(2).Value Then
            If ((txtNewPwd.Text) = (txtConPwd.Text)) Then
                RstLog.Close: RstLog.Open "SELECT * FROM tblLogin WHERE User_ID='" & frmMain.lblUserID & _
                                          "' and User_Pwd='" & txtPwd.Text & "'"
                                          
                RstLog.Fields(2).Value = txtNewPwd.Text: RstLog.Update
                MsgBox lblUserID & " Password has been Changed Successfully", vbInformation, "Change Password . . . "
                Call Ctrl_ET.Populate_Text_Clear(frmChangePwd) 'Clear TextBoxes.
                CmdCancel.Enabled = False: CmdChange.Enabled = False: CmdExit.SetFocus
                
                Unload frmChangePwd 'Unload Change Password Form.
                
            ElseIf ((txtNewPwd.Text) <> (txtConPwd.Text)) Then
                MsgBox "Choosen New Password and Confirm Password must be same." & vbCrLf & _
                       "Please try again ....... ", vbCritical, "Error! Change Password"
                       SendKeys "{Home}+{End}": txtConPwd.Text = "": txtNewPwd.SetFocus
            End If
        ElseIf txtPwd.Text <> RstLog.Fields(2).Value Then
            MsgBox "User's Password is not correct." & vbCrLf & _
                   "Please enter correct Password.", vbCritical, "Error! Incorrect Password"
                   SendKeys "{Home}+{End}": txtPwd.SetFocus
        End If
    ElseIf txtPwd.Text = "" Then
        MsgBox "Must be enter the User's Password." & vbCrLf & _
               "without User' Password you Can't Proceed.", vbCritical, "Error! User Password"
               SendKeys "{Home}+{End}": txtPwd.SetFocus
    End If
End Sub

Private Sub CmdExit_Click()
    Unload frmChangePwd
End Sub

Private Sub Form_Load()
    frmChangePwd.Move (frmMain.Width / 3), (frmMain.Height / 6)
    Call Ctrl_ET.Populate_Text_Clear(frmChangePwd) 'Call for Clearing the Text Boxes.
    lblUserID = "": lblUserName = "":    lblUserID = frmMain.lblUserID.Caption
    
    RstSQL = "SELECT * FROM tblLogin"
    RstLog.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    If lblUserID <> "" Then
        RstLog.Close: RstLog.Open "SELECT * FROM tblLogin WHERE User_ID='" & lblUserID & "'"
        lblUserName = RstLog.Fields(0).Value
        
        RstLog.Close: RstLog.Open "SELECT * FROM tblLogin"
    ElseIf lblUserID = "" Then
        Call Ctrl_ET.Populate_Entery(frmChangePwd, False) 'Not Allow Enteries(Changing Password).
    End If
    Call Ctrl_ET.Explode(frmChangePwd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstLog.Close 'Close Database.
End Sub

Private Sub txtConPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtConPwd.Text <> "" Then CmdChange.SetFocus
    End If
End Sub

Private Sub txtNewPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNewPwd.Text <> "" Then txtConPwd.SetFocus: CmdChange.Enabled = True
    End If
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPwd.Text <> "" Then txtNewPwd.SetFocus: CmdCancel.Enabled = True
    End If
End Sub
