VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCharges 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Air Line Infomation :-"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frmCharges.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton CmdSubmit 
      Height          =   330
      Left            =   1920
      TabIndex        =   20
      Top             =   6000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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
      BCOL            =   14215660
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCharges.frx":5DC3
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   4800
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   55967745
      CurrentDate     =   38961
   End
   Begin VB.TextBox txtAirCharge 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   1920
      TabIndex        =   16
      Text            =   "txtAirCharge"
      Top             =   2340
      Width           =   1095
   End
   Begin VB.TextBox txtServCharge 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   5280
      TabIndex        =   14
      Text            =   "txtServCharge"
      Top             =   2340
      Width           =   1095
   End
   Begin ctrlNSDataCombo.NSDataCombo NSDCService 
      Height          =   315
      Left            =   2760
      TabIndex        =   11
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin VB.ComboBox CmbDest 
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox CmbFrom 
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin ctrlNSDataCombo.NSDataCombo NSDCAirLine 
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
      Height          =   330
      Left            =   4680
      TabIndex        =   2
      Top             =   6000
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
      MICON           =   "frmCharges.frx":5DDF
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
      Left            =   3360
      TabIndex        =   1
      Top             =   6000
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
      MICON           =   "frmCharges.frx":5DFB
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
      Left            =   480
      TabIndex        =   0
      Top             =   6000
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
      MICON           =   "frmCharges.frx":5E17
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
   Begin MSComctlLib.ListView LstChInfo 
      Height          =   2775
      Left            =   75
      TabIndex        =   17
      Top             =   3120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Air Line"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Services"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Origin"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Destination"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Air Line"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Services"
         Object.Width           =   1764
      EndProperty
      Picture         =   "frmCharges.frx":5E33
   End
   Begin VB.Label Label3 
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
      TabIndex        =   21
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4800
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air Line Charges : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Left            =   75
      TabIndex        =   15
      Top             =   2340
      Width           =   1785
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Services Charges : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Left            =   3360
      TabIndex        =   13
      Top             =   2340
      Width           =   1935
   End
   Begin VB.Label lblServices 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Services : "
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
      Left            =   1680
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To : "
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
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From : "
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
      Left            =   2085
      TabIndex        =   4
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label lblAirLine 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Air Line : "
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
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstAL As New ADODB.Recordset 'For Opening Data base Table.(Air Lines)
Dim RstSrv As New ADODB.Recordset 'For Opening Data base Table.(Services)
Dim RstDes As New ADODB.Recordset 'For Opening Data base Table.(Destinations)
Dim RstChar As New ADODB.Recordset 'For Data Base Table.(Air Line/Services Charges)

Private Sub Pop_Save() 'Save In Database Table.
'    MsgBox "Here code for Saving in DB Table", vbOKOnly, "Underconstruction"
    With RstChar
    If Opt_Flag = "New" Then
        .AddNew
            .Fields(0).Value = lblAirLine: .Fields(1).Value = lblServices
            .Fields(2).Value = CmbFrom.Text: .Fields(3).Value = CmbDest.Text
            .Fields(4).Value = Val(txtAirCharge.Text): .Fields(5).Value = Val(txtServCharge.Text)
            .Fields(6).Value = DTPicker1.Value
        .Update: MsgBox "Record has been Saved Succesfully", _
                    vbCritical, "Record Saved": CmdCancel.Enabled = False

    ElseIf Opt_Flag = "None" Then
        If .RecordCount > 0 Then .MoveFirst
        LstChInfo.ListItems.Clear 'To Clear The List Items.
        For IntI = 1 To .RecordCount
            RstAL.Close: RstAL.Open "SELECT * FROM tblAirLine WHERE AL_ID='" & .Fields(0).Value & "'"
            RstSrv.Close: RstSrv.Open "SELECT * FROM tblServices WHERE Srv_ID='" & .Fields(0).Value & "'"
            Set LItem = LstChInfo.ListItems.Add(IntI, , RstAL.Fields(1).Value)
                LItem.SubItems(1) = RstSrv.Fields(1).Value
                LItem.SubItems(2) = .Fields(2).Value
                LItem.SubItems(3) = .Fields(3).Value
                LItem.SubItems(4) = .Fields(4).Value
                LItem.SubItems(5) = .Fields(5).Value
'                LItem.SubItems(6) = .Fields(6).Value
            If Not .EOF Then .MoveNext
            If .EOF Then Exit For
        Next
    End If
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
        If RstSrv.RecordCount <= 0 Then lblServices = ""

        .setDropWindowSize 3600, 3000
        .TextReadOnly = True: .SetDropDownTitle = "Services List"
    End With
    
    'For Air Line List In NS-Data Combo box.
    With NSDCAirLine
        .ClearColumn: .AddColumn "Code", 1200: .AddColumn "Air Line Name", 2200
        
        .Connection = DB_Conect.ConnectionString
        .sqlFields = "AL_ID, AL_Des": .sqlTables = "tblAirLine"
        .sqlSortOrder = "AL_ID ASC": .BoundField = "AL_ID"
        
        .PageBy = 25
        .DisplayCol = 2 'That Value Want to Display In Field(TextBox).
        
        RstSQL = "SELECT * FROM tblAirLine WHERE AL_Des='" & NSDCAirLine.Text & "'"
        RstAL.Close: RstAL.Open RstSQL
        If RstAL.RecordCount > 0 Then lblAirLine = RstAL.Fields(0).Value
        If RstAL.RecordCount <= 0 Then lblAirLine = ""
        
        .setDropWindowSize 3600, 3000
        .TextReadOnly = True: .SetDropDownTitle = "Air Line List"
    End With

    'For Country List In NS-Data Combo box.
'    With NSDCountry
'        .ClearColumn: .AddColumn "Country Code", 1200: .AddColumn "Country Name", 2200
'
'        .Connection = DB_Conect.ConnectionString
'        .sqlFields = "Country_ID, Country_Name": .sqlTables = "tblCountryList"
'        .sqlSortOrder = "Country_Name ASC": .BoundField = "Country_ID"
'
'        .PageBy = 25
'        .DisplayCol = 2 'That Value Want to Display In Field(TextBox).
'
'        RstSQL = "SELECT * FROM tblCountryList WHERE Country_Name='" & NSDCountry.Text & "'"
'        RstCont.Close: RstCont.Open RstSQL
'        If RstCont.RecordCount > 0 Then lblCountry = RstCont.Fields(0).Value
'        If RstCont.RecordCount <= 0 Then lblCountry = ""
'
'        .setDropWindowSize 3600, 3000
'        .TextReadOnly = True: .SetDropDownTitle = "Country List"
'    End With
End Sub

Private Sub CmbDest_GotFocus()
    Call Ctrl_ET.Populate_Init_Cmb(RstDes, 1, CmbDest) 'Call For Initilizing Combo Box.
    CmdCancel.Enabled = True
End Sub

Private Sub CmbDest_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If CmbDest.Text <> "Choose" Then txtAirCharge.SetFocus
    End If
End Sub

Private Sub CmbFrom_GotFocus()
'    Call Ctrl_Variable.Populate_LocalArea(CmbFrom) 'Call For Local Area Origin.
End Sub

Private Sub CmbFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbFrom.Text <> "Choose" Then CmbDest.SetFocus
    End If
End Sub

Private Sub CmdCancel_Click()
    Call Ctrl_ET.Populate_Text_Clear(frmCharges) 'Initialize the Form.
    Opt_Flag = "None": Call Pop_Save 'For Operating Record.
    Call Ctrl_ET.Populate_Entery(frmCharges, False) 'Ristrict the User to Entery n Text Boxes.
    CmdNew.Enabled = True: CmdCancel.Enabled = False: CmdSubmit.Enabled = False
    CmdNew.SetFocus: Opt_Flag = "None"
End Sub

Private Sub CmdExit_Click()
    Unload frmCharges
End Sub

Private Sub CmdNew_Click()
    Call Ctrl_ET.Populate_Entery(frmCharges, True) 'Allow the User to Entery n Text Boxes.
    CmdNew.Enabled = Not CmdNew.Enabled
    Opt_Flag = "New": NSDCAirLine.SetFocus
End Sub

Private Sub CmdSubmit_Click()
    Call Pop_Save 'For Operating Record.
    Call Ctrl_ET.Populate_Text_Clear(frmCharges) 'To Clear the Text Boxes.
    Call Ctrl_ET.Populate_Entery(frmCharges, False) 'Ristrict the User to Entery n Text Boxes.
    Opt_Flag = "None": Call Pop_Save 'For Operating Record.
    CmdNew.Enabled = True: CmdCancel.Enabled = False: CmdSubmit.Enabled = False
    CmdNew.SetFocus
End Sub

Private Sub Form_Load()
    frmCharges.Move (frmMain.Width / 3), (frmMain.Height / 8)
    Call Ctrl_ET.Populate_Text_Clear(frmCharges) 'Call For clearing Txts
    Opt_Flag = "None": Call Ctrl_Variable.Populate_LocalArea(CmbFrom) 'Call For Local Area Origin.
    DTPicker1.Value = Date 'Assign the Today Date.
    
    RstSQL = "Select * From tblAirLine" 'Opening tblAirLine Database Table.
    RstAL.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "Select * From tblServices" 'Opening tblAirLine Database Table.
    RstSrv.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "Select * From tblDestination" 'Opening tblDestination Database Table.
    RstDes.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    RstSQL = "Select * From tblCharges" 'Opening tblCharges Database Table.
    RstChar.Open RstSQL, DB_Conect, adOpenStatic, adLockOptimistic
    
    Call Ctrl_ET.Populate_Init_Cmb(RstDes, 1, CmbFrom)
    Call Ctrl_ET.Populate_Init_Cmb(RstDes, 1, CmbDest)
    Call InitNSD 'For Initialize Combo Box.
    
    Call Pop_Save 'For Operating Record.
    Call Ctrl_ET.Populate_Entery(frmCharges, False) 'Ristrict the User to Entery n Text Boxes.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstAL.Close 'Closing DB Table.
    RstSrv.Close 'Closing DB Table.
    RstDes.Close 'Closing DB Table.
    RstChar.Close 'Closing DB Table.
End Sub

Private Sub NSDCAirLine_Change()
    Call InitNSD 'Call For NSDCServices Text Box.
End Sub

Private Sub NSDCService_Change()
    Call InitNSD 'Call For NSDCServices Text Box.
End Sub

Private Sub txtAirCharge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtAirCharge.Text <> "" Then txtServCharge.SetFocus: CmdSubmit.Enabled = True
    End If
End Sub

Private Sub txtServCharge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtServCharge.Text <> "" Then CmdSubmit.SetFocus
    End If
End Sub
