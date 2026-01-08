Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub dept_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartStepTemplateApprovals", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")
Form_sfrmPartProjectTemplate.approvalCount.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If IsNull(Me.recordId) Then
    MsgBox "This is an empty record, don't worry about deleting this. It's for creating new approvals", vbInformation, "Can't do that"
    Exit Sub
End If

Call registerWdbUpdates("tblPartStepTemplateApprovals", Me.recordId, "DELETE", Me.dept, "DELETE", "frmPartProjectTemplate")
dbExecute ("DELETE FROM tblPartStepTemplateApprovals WHERE [recordId] = " & Me.recordId)

Me.Requery

MsgBox "Approval Deleted", vbOKOnly, "Deleted"

Form_sfrmPartProjectTemplate.approvalCount.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub reqLevel_AfterUpdate()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
Call validate
Form_sfrmPartProjectTemplate.approvalCount.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Function validate()

validate = True

If (Me.dept = "" Or Me.reqLevel = "") Then
    MsgBox "please enter department and level", vbCritical, "Error"
    validate = False
End If

End Function
