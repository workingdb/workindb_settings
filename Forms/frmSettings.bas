Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim dbLoc, FSO

Function checkIfAdminDev() As Boolean

checkIfAdminDev = False

Dim errorTxt As String: errorTxt = ""
If (privilege("admin") = False) Then errorTxt = "You need admin privilege to do this"
If (privilege("developer") = False) Then errorTxt = "You need developer privilege to do this"

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Access Denied"
    checkIfAdminDev = False
    Exit Function
End If

checkIfAdminDev = True

End Function

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub cpcHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = 'sfrmCPC_templates'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub disShift_Click()
On Error GoTo Err_Handler
If Not checkIfAdminDev Then Exit Sub
ap_DisableShift
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub enableShift_Click()
On Error GoTo Err_Handler
If Not checkIfAdminDev Then Exit Sub
ap_EnableShift
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub emailUsers_Click()
On Error GoTo Err_Handler

If Not checkIfAdminDev Then Exit Sub

Dim strTo As String
Dim SendItems As New clsOutlookCreateItem

    Dim db As Database
    Dim rs1 As Recordset, rsFiltered As Recordset
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPermissions", dbOpenSnapshot)
    rs1.filter = "[Inactive] = false"
    Set rsFiltered = rs1.OpenRecordset
    strTo = ""

    Dim lngCnt As Long
    lngCnt = 0
    Do While Not rsFiltered.EOF
        strTo = strTo & getEmail(rsFiltered![User]) & "; "
        lngCnt = lngCnt + 1
        rsFiltered.MoveNext
    Loop

    rs1.Close
    Set rs1 = Nothing
    rsFiltered.Close
    Set rsFiltered = Nothing
    Set db = Nothing

SendItems.CreateMailItem sendTo:="", _
                             BCC:=strTo, _
                             subject:="Working DB"
    Set SendItems = Nothing
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub hideNav_Click()
On Error GoTo Err_Handler
Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
Call DoCmd.RunCommand(acCmdWindowHide)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub hideRibbon_Click()
On Error GoTo Err_Handler

DoCmd.ShowToolbar "Ribbon", acToolbarNo

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = 'frmPartTrackingSettings'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub openMeetingChecklists_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartMeetingChecklists"
Form_frmPartMeetingChecklists.filter = "recordId = 0"
Form_frmPartMeetingChecklists.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub showRibbon_Click()
On Error GoTo Err_Handler

DoCmd.ShowToolbar "Ribbon", acToolbarYes

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub showNav_Click()
On Error GoTo Err_Handler

Call DoCmd.SelectObject(acTable, , True)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

If privilege("Edit") = True Then
    Me.plmSettings.Enabled = True
Else
    Me.plmSettings.Enabled = False: Me.plmSettings.Caption = "Edit Privilege Needed"
End If

Me.naviTab.Pages(0).Visible = privilege("Developer") 'dev dashboard
Me.naviTab.Pages(1).Visible = privilege("Admin") 'permissions
Me.naviTab.Pages(2).Visible = privilege("Edit") 'links
Me.naviTab.Pages(3).Visible = privilege("Training_Mode") 'dropdowns
Me.naviTab.Pages(4).Visible = Not restrict(Environ("username"), "Project", "Supervisor", True) Or privilege("Developer") Or Not restrict(Environ("username"), "Service", "Supervisor") 'part tracking
Me.naviTab.Pages(5).Visible = privilege("Edit") 'holidays

If CBool(userData("Developer")) Then
    Me.lblFrmProjectTemplates.Caption = "All Project Templates"
    Me.sfrmProjectTemplateSettings.Form.FilterOn = False
    Form_sfrmProjectTemplateSettings.AllowAdditions = False
    GoTo afterFiltTemps
End If

If Not restrict(Environ("username"), "Service", "Supervisor") Then
    Me.lblFrmProjectTemplates.Caption = "Service Project Templates"
    Me.sfrmProjectTemplateSettings.Form.filter = "templateType = 2"
    Form_sfrmProjectTemplateSettings.templateType.DefaultValue = 2
    Me.sfrmProjectTemplateSettings.Form.FilterOn = True
ElseIf Not restrict(Environ("username"), "Project", "Supervisor") Then
    Me.lblFrmProjectTemplates.Caption = "New Model Project Templates"
    Me.sfrmProjectTemplateSettings.Form.filter = "templateType = 1"
    Me.sfrmProjectTemplateSettings.Form.FilterOn = True
    Form_sfrmProjectTemplateSettings.templateType.DefaultValue = 1
Else
    Me.sfrmProjectTemplateSettings.Visible = False
End If

afterFiltTemps:
'code for dev dashboard:
Me.filter = "[ID] = 1"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub plmSettings_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPLMsettings"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub themeEditor_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmThemeEditor"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub viewProjectProjects_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartProjects"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
