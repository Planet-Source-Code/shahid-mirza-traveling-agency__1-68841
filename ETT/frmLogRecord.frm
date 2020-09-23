VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmLogRecord 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " User's Login Record . . ."
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10755
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
   ScaleHeight     =   8835
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   8760
      Top             =   8400
   End
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   360
      Left            =   9240
      TabIndex        =   2
      Top             =   8460
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
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
      MICON           =   "frmLogRecord.frx":0000
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
   Begin MSComctlLib.ListView LstLogRecord 
      Height          =   8060
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   14208
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Login Date"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Login Time"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "LogOut Date"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "LogOut Time"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   1764
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
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   8520
      Width           =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Login Records : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   8475
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   " List Of Users Log In Record . . . "
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
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10845
   End
End
Attribute VB_Name = "frmLogRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstLogRecord As New ADODB.Recordset
Dim RstLogIn As New ADODB.Recordset

Private Sub LogIn_Record()
    With RstLogRecord
        LstLogRecord.ListItems.Clear 'Clearing List
        If .RecordCount > 0 Then
            For IntI = 1 To .RecordCount
                RstLogIn.Close: RstLogIn.Open "SELECT * FROM tblLogIn WHERE User_ID='" & LogIn_UID & "'"
                Set LItem = LstLogRecord.ListItems.Add(IntI, , RstLogIn.Fields(0).Value)
                LItem.SubItems(1) = .Fields(0)
                LItem.SubItems(2) = .Fields(1)
                LItem.SubItems(3) = .Fields(2)
                If IsNull(.Fields(3).Value) = False Then LItem.SubItems(4) = .Fields(3)
                If IsNull(.Fields(4).Value) = False Then LItem.SubItems(5) = .Fields(4)
                If .EOF = False Then .MoveNext
                If .EOF = True Then Exit For
            Next
            Label2.Caption = Label2.Caption & " " & LstLogRecord.ListItems.Count
        ElseIf .RecordCount <= 0 Then
            MsgBox "There is not record in Login Record File.", vbCritical, "Error!Record Not Found"
        End If
    End With
    Timer1.Interval = 0
End Sub

Private Sub CmdExit_Click()
    Unload frmLogRecord
End Sub

Private Sub Form_Load()
    frmLogRecord.Move (frmMain.Width / 8), (frmMain.Height / 100)
    
    RstSQL = "SELECT * FROM tblLogIn_Record"
    RstLogRecord.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic

    RstSQL = "SELECT * FROM tblLogIn"
    RstLogIn.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    Timer1.Interval = 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstLogIn.Close: RstLogRecord.Close
End Sub

Private Sub Timer1_Timer()
    Call LogIn_Record 'Function/Procedure Call
End Sub
