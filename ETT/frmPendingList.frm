VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmPendingList 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7365
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
   Picture         =   "frmPendingList.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   7080
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
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPendingList.frx":76242
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
   Begin LVbuttons.LaVolpeButton CmdViewRec 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&View Record"
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
      MICON           =   "frmPendingList.frx":7625E
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
   Begin VB.OptionButton Opt_NIC 
      BackColor       =   &H80000005&
      Caption         =   "NIC #"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5160
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Opt_Passport 
      BackColor       =   &H80000005&
      Caption         =   "Passport #"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Text            =   "txtNo"
      Top             =   1080
      Width           =   2895
   End
   Begin MSComctlLib.ListView LstPending 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9340
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   13830917
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr #"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Passport #"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NIC #"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Submit Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Return Date"
         Object.Width           =   2469
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
      ForeColor       =   &H00C0C0C0&
      Height          =   450
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   2400
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   1545
      Width           =   7095
   End
   Begin VB.Label lblTextNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Passport # : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " List of Pending "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   2385
   End
End
Attribute VB_Name = "frmPendingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstPend As New ADODB.Recordset

Private Sub Pendig_List_Generate()
    With RstPend
        For IntI = 1 To .RecordCount
            Set LItem = LstPending.ListItems.Add(IntI, , IntI)
                LItem.SubItems(1) = .Fields(1).Value
                LItem.SubItems(2) = .Fields(2).Value
                LItem.SubItems(3) = .Fields(8).Value
                LItem.SubItems(4) = .Fields(9).Value
                If .EOF = False Then .MoveNext
                If .EOF = True Then Exit Sub
        Next
    End With
End Sub

Private Sub CmdExit_Click()
    Unload frmPendingList
End Sub

Private Sub CmdViewRec_Click()

    Catch_Field = LstPending.SelectedItem
'    MsgBox Catch_Field
    For IntI = 1 To LstPending.ListItems.Count
        Set LItem = LstPending.ListItems.Item(IntI)
            If LItem = Catch_Field Then
                Passport = LItem.SubItems(1) 'Passport Number Move to Variable
                CNIC = LItem.SubItems(2) 'CNIC Number Move to Variable
                Exit For
            End If
    Next
    Load frmCheckStatus: frmCheckStatus.Show
End Sub

Private Sub Form_Load()
    frmPendingList.Move (frmMain.Width / 3), (frmMain.Height / 14)
    Call Ctrl_ET.Populate_Text_Clear(frmPendingList) 'To Clearing the Textboxes.
    txtNo.Enabled = False
    
    RstSQL = "SELECT * FROM tblCustomer_Info"
    RstPend.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstPend.Close: RstPend.Open "SELECT * FROM tblCustomer_Info WHERE Status_Job=" & False
    Call Pendig_List_Generate 'For Generating the List Records.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstPend.Close 'To Close the DB Table.
End Sub

Private Sub LstPending_Click()
    CmdViewRec.Enabled = True
End Sub

Private Sub LstPending_DblClick()
    Call CmdViewRec_Click 'Call Command Button for Active its Function.
End Sub

Private Sub Opt_NIC_Click()
    If Opt_NIC.Value = True Then
        lblTextNo.Caption = "Enter NIC # : "
        txtNo.Enabled = True: SendKeys "{Home}+{End}": txtNo.SetFocus
    ElseIf Opt_NIC.Value = False Then
        txtNo.Enabled = False
    End If
End Sub

Private Sub Opt_Passport_Click()
    If Opt_Passport.Value = True Then
        lblTextNo.Caption = "Enter Passport # : "
        txtNo.Enabled = True: SendKeys "{Home}+{End}": txtNo.SetFocus
    ElseIf Opt_Passport.Value = False Then
        txtNo.Enabled = False
    End If
End Sub

Private Sub txtNo_Change()
    With LstPending
        If txtNo.Text <> "" Then
            For IntI = 1 To .ListItems.Count
                Set LItem = .ListItems.Item(IntI)
                    If Opt_Passport.Value = True Then
                        If LItem.SubItems(1) = txtNo.Text Then
                            
                        End If
                    ElseIf Opt_NIC.Value = True Then
                        If LItem.SubItems(2) = txtNo.Text Then
                        
                        End If
                    End If
            Next
        End If
    End With
End Sub
