VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmServices 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Services Infomation :-"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frmServices.frx":0000
   ScaleHeight     =   5340
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   330
      Left            =   3720
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmServices.frx":52E4
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
      Height          =   330
      Left            =   2400
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      BCOL            =   14215660
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmServices.frx":5300
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
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   330
      Left            =   840
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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
      BCOL            =   14215660
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmServices.frx":531C
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
   Begin MSComctlLib.ListView LstSerInfo 
      Height          =   2775
      Left            =   840
      TabIndex        =   7
      Top             =   2040
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
      Picture         =   "frmServices.frx":5338
   End
   Begin VB.TextBox txtSerDes 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Text            =   "txtSerDes"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtServID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Text            =   "txtServID"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "+ 9 2 3 0 0 6 4 1 0 7 5 8"
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
      Height          =   2655
      Left            =   5400
      TabIndex        =   10
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "M  S H A H I D  A S L A M  M U G H A L"
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
      Height          =   4290
      Left            =   5160
      TabIndex        =   9
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Existing Informations List "
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
      Left            =   960
      TabIndex        =   8
      Top             =   1770
      Width           =   2640
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   300
      Left            =   840
      Top             =   1725
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Service ID : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstSrv As New ADODB.Recordset 'For Opening Data base Table.

Private Sub Pop_Save() 'Save In Database Table.
'    MsgBox "Here code for Saving in DB Table", vbOKOnly, "Underconstruction"
    With RstSrv
    If Opt_Flag = "New" Then
        .AddNew
            .Fields(0).Value = txtServID.Text
            .Fields(1).Value = txtSerDes.Text
        .Update: MsgBox "Record has been Saved Succesfully", _
                    vbCritical, "Record Saved": CmdCancel.Enabled = False
        CmdNew.Enabled = True: txtServID.Text = "Auto ID": txtSerDes.Text = ""
        
    ElseIf Opt_Flag = "None" Then
        If .RecordCount > 0 Then .MoveFirst
        LstSerInfo.ListItems.Clear 'To Clear The List Items.
        For IntI = 1 To .RecordCount
            Set LItem = LstSerInfo.ListItems.Add(IntI, , .Fields(0).Value)
                LItem.SubItems(1) = .Fields(1).Value
            If Not .EOF Then .MoveNext
            If .EOF Then Exit For
        Next
    End If
    End With
End Sub

Private Sub CmdCancel_Click()
    CmdNew.Enabled = True: CmdCancel.Enabled = False
    txtServID.Text = "Auto ID": txtSerDes.Text = ""
    CmdNew.SetFocus: Opt_Flag = "None"
    Call Pop_Save 'For Operating Record.
End Sub

Private Sub CmdExit_Click()
    Unload frmServices
End Sub

Private Sub CmdNew_Click()
    CmdNew.Enabled = Not CmdNew.Enabled
    txtSerDes.SetFocus: Opt_Flag = "New"
End Sub

Private Sub Form_Load()
    frmServices.Move (frmMain.Width / 3), (frmMain.Height / 8)
    Call Ctrl_ET.Populate_Text_Clear(frmServices) 'Call For clearing Txts
    txtServID.Text = "Auto ID": Opt_Flag = "None"
    
    RstSQL = "Select * From tblServices" 'Opening tblServices Database Table.
    RstSrv.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    Call Pop_Save 'For Operating Record.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstSrv.Close 'Closing DB Tables.
End Sub

Private Sub txtSerDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Ctrl_ET.Populate_CheckList(LstSerInfo, txtSerDes) 'To aviod the Dublication in List Items.
        If Find_Flag = False Then
            Call Pop_Save 'To Saving In Database Table.
            Opt_Flag = "None": Call Pop_Save 'For Operating Record.
            CmdNew.SetFocus 'Focus Set To New Button.
        End If
    Else
        If txtSerDes.Text <> "" Then CmdCancel.Enabled = True
        'Call Following Function For First Capital Latter.
        Call Ctrl_ET.Populate_Alpha_Char(KeyAscii, frmServices, txtSerDes)
        Call Ctrl_ET.Populate_AutoID(RstSrv) 'Calling For Auto ID Generation.
        txtServID.Text = Auto_ID 'Putting Auto ID In Text Box.
    End If
End Sub
