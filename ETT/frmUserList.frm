VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmUserList 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " User's List . . ."
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5790
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
   ScaleHeight     =   7380
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton CmdCheck_Status 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Check Status"
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
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmUserList.frx":0000
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
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
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
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmUserList.frx":001C
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
   Begin MSComctlLib.ListView LstLoginUser 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11456
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   2293
      EndProperty
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
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   " List Of Log In Users . . . "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   -75
      TabIndex        =   2
      Top             =   0
      Width           =   5955
   End
End
Attribute VB_Name = "frmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstLogIn As New ADODB.Recordset
Dim LItem As ListItem

Private Sub View_Rec()
    Catch_Field = LstLoginUser.SelectedItem
'    MsgBox Catch_Field
    For IntI = 1 To LstLoginUser.ListItems.Count
        Set LItem = LstLoginUser.ListItems.Item(IntI)
            If LItem = Catch_Field Then
                Passport = LItem.SubItems(1) 'Passport Number Move to Variable
                CNIC = LItem.SubItems(2) 'CNIC Number Move to Variable
                Exit For
            End If
    Next
    Load frmCheckStatus: frmCheckStatus.Show
End Sub

Private Sub LogIn_Record()
    With RstLogIn
        If .RecordCount > 0 Then
            For IntI = 1 To .RecordCount
                Set LItem = LstLoginUser.ListItems.Add(IntI, , .Fields(0).Value)
                    LItem.SubItems(1) = .Fields(1).Value
                    LItem.SubItems(2) = .Fields(3).Value
                    If .EOF = True Then Exit For
                    If .EOF = False Then .MoveNext
            Next
        ElseIf .RecordCount <= 0 Then
            MsgBox "There is not record in Login Record File.", vbCritical, "Error!Record Not Found"
        End If
    End With
End Sub

Private Sub CmdExit_Click()
    Unload frmUserList
End Sub

Private Sub Form_Load()
    frmUserList.Move (frmMain.Width / 2.5), (frmMain.Height / 12)
    RstSQL = "SELECT * FROM tblLogIn"
    RstLogIn.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    Call LogIn_Record 'Function/Procedure Call
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstLogIn.Close
End Sub
