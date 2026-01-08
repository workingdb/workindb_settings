Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Dim gateTemplateId
gateTemplateId = DMin("recordId", "tblPartGateTemplate", "projectTemplateId = " & Me.recordId)

On Error GoTo invis
Me.sfrmPartProjectTemplate.Form.filter = "gateTemplateId = " & gateTemplateId
Me.sfrmPartProjectTemplate.Form.gateTemplateId.DefaultValue = gateTemplateId
Me.sfrmPartProjectTemplate.Form.FilterOn = True

Exit Sub
invis:
Me.sfrmPartProjectTemplate.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = '" & Me.name & "'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler

Me.Requery
Me.sfrmPartProjectTemplate.Requery
Me.sfrmPartProjectTemplateGates.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
