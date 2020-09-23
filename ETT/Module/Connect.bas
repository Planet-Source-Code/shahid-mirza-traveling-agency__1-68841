Attribute VB_Name = "Connect"
Public Sub Main()
    DB_Conect.Open "DSN=Travel Agency", Admin, , -1

    Load frmEID: frmEID.Show
    frmMain.lblUserID = "": frmMain.lblDate = "": frmMain.lblTime = ""
End Sub

Public Sub User_Privilliage(UID As String)
    Dim RstUserPrivilage As New ADODB.Recordset
    RstUserPrivilage.Open "SELECT * FROM tblLogIn", DB_Conect, adOpenStatic, adLockOptimistic
    
    RstUserPrivilage.Close: RstUserPrivilage.Open "SELECT * FROM tblLogIn WHERE User_ID='" & UID & "'"
    If RstUserPrivilage.RecordCount > 0 Then
        If RstUserPrivilage.Fields(3).Value = "Active" Then
            With frmMain
                If RstUserPrivilage.Fields(4).Value = "Administrator" Then
                    Load frmMain: frmMain.Show: frmMain.Picture = LoadPicture(App.Path & "\Images\MainBG1024x768.JPG")
                    frmMain.Toolbar1.Visible = True: Load frmSplash: frmSplash.Show: Load frmJobRpt: frmJobRpt.Show
                    
                ElseIf RstUserPrivilage.Fields(4).Value = "Limited User" Then
                    .lblUserList.Enabled = False: .lblLogDetail.Enabled = False: .lblResetDB.Enabled = False 'Admin Only.
                    .lblAirLine.Enabled = False: .lblDestination.Enabled = False: .lblServices.Enabled = False: .lblCharges.Enabled = False 'Navigation
                    .lblNewUser.Enabled = False 'Security.
                    
                    Load frmMain: frmMain.Show: frmMain.Picture = LoadPicture(App.Path & "\Images\MainBG1024x768.JPG")
                    frmMain.Toolbar1.Visible = True: Load frmSplash: frmSplash.Show: Load frmJobRpt: frmJobRpt.Show
                    
                ElseIf RstUserPrivilage.Fields(4).Value = "Guest" Then
                    .lblUserList.Enabled = False: .lblLogDetail.Enabled = False: .lblResetDB.Enabled = False 'Admin Only.
                    .lblNewJob.Enabled = False: .lblCheckStatus.Enabled = False: .lblAirLine.Enabled = False: .lblDestination.Enabled = False: .lblServices.Enabled = False: .lblCharges.Enabled = False 'Navigation
                    .lblJobpending_Rpt.Enabled = False: .lblDoneJob_Rpt.Enabled = False: .lblDaily_Month_Rpt.Enabled = False 'Reports.
                    .lblNewUser.Enabled = False: .lblChangePwd.Enabled = False 'Security.
                    Load frmMain: frmMain.Show: frmMain.Picture = LoadPicture(App.Path & "\Images\MainBG1024x768.JPG")
                    Load frmSplash: frmSplash.Show
                    
                End If
            End With
        Else
            MsgBox "Sorry ! " & vbCrLf & "This not allowed to Access the System" & vbCrLf & _
                   "Please: Contact the Administrator", vbCritical, "Error ! Unauthorized Access"
                   End
        End If
    End If
End Sub
