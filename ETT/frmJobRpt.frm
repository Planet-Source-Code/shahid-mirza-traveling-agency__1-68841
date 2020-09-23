VERSION 5.00
Begin VB.Form frmJobRpt 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4185
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
   ScaleHeight     =   435
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblPending 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pending : "
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
      Left            =   0
      MouseIcon       =   "frmJobRpt.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmJobRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstCust As New ADODB.Recordset

Private Sub Form_Load()
    frmJobRpt.Move (frmMain.Width / 20), (frmMain.Height / 1.3)
    
    RstSQL = "SELECT * FROM tblCustomer_Info"
    RstCust.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstCust.Close: RstCust.Open "SELECT * FROM tblCustomer_Info WHERE Status_Job=" & False
    lblPending = "Pending : " & RstCust.RecordCount & "."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstCust.Close 'Closing Database.
End Sub

Private Sub lblPending_Click()
    Load frmPendingList: frmPendingList.Show
End Sub
