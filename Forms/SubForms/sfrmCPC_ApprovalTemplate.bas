Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub dept_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblCPC_Approvals_Template", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "sfrmCPC_ApprovalTemplate")
Form_sfrmCPC_StepTemplate.approvalCount.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If IsNull(Me.ID) Then
    MsgBox "This is an empty record, don't worry about deleting this. It's for creating new approvals", vbInformation, "Can't do that"
    Exit Sub
End If

Call registerWdbUpdates("tblCPC_Approvals_Template", Me.ID, "DELETE", Me.dept, "DELETE", "sfrmCPC_ApprovalTemplate")
dbExecute ("DELETE FROM tblCPC_Approvals_Template WHERE [ID] = " & Me.ID)

Me.Requery

MsgBox "Approval Deleted", vbOKOnly, "Deleted"

Form_sfrmCPC_StepTemplate.approvalCount.Requery

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
Form_sfrmCPC_StepTemplate.approvalCount.Requery

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
