Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.txtCF = Me.recordId

If IsNull(Me.recordId) = False Then
    Form_sfrmPartProjectTemplate.Visible = True
    Form_sfrmPartProjectTemplate.filter = "gateTemplateId = " & Me.recordId
    Form_sfrmPartProjectTemplate.gateTemplateId.DefaultValue = Me.recordId
    Form_sfrmPartProjectTemplate.FilterOn = True
Else
    Form_sfrmPartProjectTemplate.Visible = False
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub gateDuration_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartGateTemplate", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub gateTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartGateTemplate", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
