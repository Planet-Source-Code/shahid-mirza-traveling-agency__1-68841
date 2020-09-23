VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4515
   ClientLeft      =   3555
   ClientTop       =   1920
   ClientWidth     =   6855
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
   Moveable        =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Chk_Unmask 
      BackColor       =   &H8000000E&
      Caption         =   "Unmasked Password"
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
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Top             =   2760
      Width           =   2655
   End
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   4035
      Width           =   1575
      _ExtentX        =   2778
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmLogin.frx":54C1
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
   Begin LVbuttons.LaVolpeButton CmdConnect 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   4035
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Connect"
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmLogin.frx":54DD
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
   Begin VB.TextBox txtPassword 
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
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "txtPassword"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ComboBox CmbUserID 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
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
      Height          =   450
      Left            =   0
      TabIndex        =   13
      Top             =   3840
      Width           =   2400
   End
   Begin VB.Label lblTry 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1800
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblTries 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Remaining Tries : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2120
      TabIndex        =   10
      Top             =   3140
      Width           =   660
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking the Unmasked Password (Upper Option) will retrive the entered Password in Alpha Numeric Character."
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   3140
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Conneting to System "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1425
      TabIndex        =   5
      Top             =   0
      Width           =   5430
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password : "
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
      Left            =   1530
      TabIndex        =   2
      Top             =   2280
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select User Name : "
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
      Left            =   1530
      TabIndex        =   1
      Top             =   1800
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":54F9
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2120
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstLog As New ADODB.Recordset
Dim RstLogRec As New ADODB.Recordset
Dim Opt_End As String

Private Sub CmbUserID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbUserID.Text <> "Choose" Then
            txtPassword.SetFocus: CmdConnect.Enabled = True
        End If
    End If
End Sub

Private Sub CmdConnect_Click()
    If txtPassword.Text <> "" Then
        RstLog.Close: RstLog.Open "SELECT * FROM tblLogin WHERE User_ID='" & Trim(CmbUserID.Text) & "'"
            If txtPassword.Text = RstLog.Fields(2).Value Then 'If User & Password Matches.
            
                 LogIn_UID = CmbUserID.Text 'Asign User ID.
                 UserName = RstLog.Fields(0).Value 'Assign User Name.
                 
                    Call Connect.User_Privilliage(CmbUserID.Text) 'Call 4 User Privilliages.
                    
                 With RstLogRec 'Enter the Record in Record Login Table.
                    .AddNew
                        .Fields(0).Value = LogIn_UID: .Fields(1).Value = Date: .Fields(2).Value = Time
                    .Update
                    .MoveLast: LogIn_Time = .Fields(2).Value
                 End With
                Call Ctrl_ET.Deplode(frmLogin) 'call 4 deploide the form untill unload form.
                
            ElseIf txtPassword.Text <> RstLog.Fields(2).Value Then 'If User & Password Not Matches.
                
                lblTries.Visible = True: lblTry.Visible = True: lblTry = Val(lblTry) - 1
                MsgBox "Password Unmatched. Becareful and Try again" & vbCrLf & _
                    "Now you have Only " & lblTry & " Tries", vbCritical, "Password Error!"
                SendKeys "{Home}+{End}": txtPassword.SetFocus
                    
                If Val(lblTry) = 0 Then MsgBox "System has blocked your ID." & vbCrLf & _
                    "Please contact your Administrator", vbCritical, "Error! User ID & Password": _
                    Call Ctrl_ET.Deplode(frmLogin): End
            End If
    ElseIf txtPassword.Text = "" Then
        MsgBox "Dont' remain Enpty." & vbCrLf & _
            "Must be enter Password.", vbCritical, "Error! Password": SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub CmdExit_Click()
    Opt_End = MsgBox("Sure you Shuting Down the System." & vbCrLf & _
                     "Please Verify. . . . . . . . . .", vbYesNo + vbCritical, "Shut Down")
    If Opt_End = vbYes Then
        Call Ctrl_ET.Deplode(frmLogin): End
    ElseIf Opt_End = vbNo Then
        Call Ctrl_ET.Populate_Text_Clear(frmLogin) 'Call for Clearing the Text Boxes.
        CmdConnect.Enabled = True: Chk_Unmask.Value = Unchecked
        CmbUserName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    frmLogin.Move (frmMain.Width / 4), (frmMain.Height / 100)
    Call Ctrl_ET.Populate_Text_Clear(frmLogin) 'Call for Clearing the Text Boxes.
    
    RstSQL = "SELECT * FROM tblLogin" 'For Login the System.
    RstLog.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "SELECT * FROM tblLogIn_Record" 'For Login Records.
    RstLogRec.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    Call Ctrl_ET.Populate_Init_Cmb(RstLog, 1, CmbUserID) 'Initiate the Combo Box.
    
    frmMain.Picture = LoadPicture(""): frmMain.Toolbar1.Visible = False
    Call Ctrl_ET.Explode(frmLogin)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstLog.Close: RstLogRec.Close
    
    With frmMain
        .lblUserID = LogIn_UID 'User ID Time & Date On Main Form.
        .lblDate = Date:  .lblTime = Time
    End With
End Sub

Private Sub txtPassword_GotFocus()
    If CmbUserID.Text <> "Choose" Then
        txtPassword.SetFocus: CmdConnect.Enabled = True
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPassword.Text <> "" Then CmdConnect.SetFocus: Call CmdConnect_Click 'Call For Function
    End If
End Sub
