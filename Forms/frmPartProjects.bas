Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub details_Click()
On Error GoTo Err_Handler

Dim errorTxt As String
errorTxt = ""

Select Case DLookup("templateType", "tblPartProjectTemplate", "recordId = " & Me.projectTemplateId)
    Case 1 'project
        If restrict(Environ("username"), "Project", "Supervisor", True) Then errorTxt = "Only Project Supervisors/Managers can open this type of project"
    Case 2 'service
        If restrict(Environ("username"), "Service", "Supervisor", True) Then errorTxt = "Only Service Supervisors/Managers can open this type of project"
End Select

If errorTxt <> "" And userData("Developer") = False Then
    MsgBox errorTxt, vbCritical, "..."
    Exit Sub
End If

DoCmd.OpenForm "frmPartProjectActions", , , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub lblPN_Click()
    Me.partNumber.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblRecordId_Click()
    Me.recordId.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblStartDate_Click()
    Me.projectStartDate.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblStatus_Click()
    Me.partProjectStatus.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblTemplate_Click()
    Me.projectTitle.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblTemplateType_Click()
    Me.templateType.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
