Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
On Error Resume Next

If IsNull(Me.ID) = False Then
    Form_frmSettings.sfrmCPC_StepTemplate.Visible = True
    Form_frmSettings.sfrmCPC_StepTemplate.Form.filter = "projectTemplateID = " & Me.ID
    Form_frmSettings.sfrmCPC_StepTemplate.Form.Controls("projectTemplateID").DefaultValue = Me.ID
    Form_frmSettings.sfrmCPC_StepTemplate.Form.FilterOn = True
Else
    Form_frmSettings.sfrmCPC_StepTemplate.Visible = False
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub projectTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblCPC_Project", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "sfrmCPC_templates")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
