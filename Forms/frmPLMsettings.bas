Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Ctl3DexSelect_AfterUpdate()
On Error GoTo Err_Handler

Me.sfrmPLMdropDowns.Form.lblDropdown.Caption = Me.Ctl3DexSelect.Value
Me.sfrmPLMdropDowns.Form.values.ControlSource = Me.Ctl3DexSelect.Value

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim adminPriv, editpriv As Boolean
Me.history.SetFocus


adminPriv = privilege("Admin")
editpriv = privilege("Edit")

Me.tabs.Pages(0).Visible = False
Me.tabs.Pages(1).Visible = False
Me.tabs.Pages(2).Visible = False
Me.tabs.Pages(3).Visible = False
Me.tabs.Pages(4).Visible = False
Me.tabs.Pages(5).Visible = False

If editpriv = True Then
    Me.tabs.Pages(3).Visible = True
    Me.tabs.Pages(4).Visible = True
    Me.tabs.Pages(5).Visible = True
Else
    MsgBox "You need Edit priviledges to access this page", vbCritical, "Access Denied"
    Form_frmPLMsettings.SetFocus
    DoCmd.Close
    Exit Sub
End If

If adminPriv = True Then
    Me.tabs.Pages(0).Visible = True
    Me.tabs.Pages(1).Visible = True
    Me.tabs.Pages(2).Visible = True
End If

Me.sfrmPLMmaterialGrades.Form.Material_Type_ID.ForeColor = vbWhite

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub matFilter_Click()
On Error GoTo Err_Handler

Call Me.sfrmPLMmaterialTypes.Form.Material_Type_Click
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler

Me.Requery
Me.sfrmPLMdropDowns.Form.Requery
Me.sfrmPLMproperties.Form.Requery
Me.sfrmPLMsection.Form.Requery
Me.sfrmPLMsettings.Form.Requery
Me.sfrmPLMmaterialTypes.Form.Requery
Me.sfrmPLMmaterialGrades.Form.Requery
Me.sfrmPLMmaterialSpecs.Form.Requery
Me.sfrmPLMmaterialNum.Form.Requery
Me.sfrmPLMcustomers.Form.Requery
Me.refresh
Me.sfrmPLMdropDowns.Form.refresh
Me.sfrmPLMproperties.Form.refresh
Me.sfrmPLMsection.Form.refresh
Me.sfrmPLMsettings.Form.refresh
Me.sfrmPLMmaterialTypes.Form.refresh
Me.sfrmPLMmaterialGrades.Form.refresh
Me.sfrmPLMmaterialSpecs.Form.refresh
Me.sfrmPLMmaterialNum.Form.refresh
Me.sfrmPLMcustomers.Form.refresh
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub requestTypeSelect_AfterUpdate()
On Error GoTo Err_Handler

Me.sfrmDesignWOtasks.Form.filter = "[Title] = '" & Me.requestTypeSelect.Value & "'"
Me.sfrmDesignWOtasks.Form.FilterOn = True
Me.sfrmDesignWOtasks.Form.Controls("taskTitle").DefaultValue = "'" & Me.requestTypeSelect.Value & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub sfrmPLMmaterialGrades_Enter()
Me.sfrmPLMmaterialGrades.Form.Requery
Me.sfrmPLMmaterialGrades.Form.refresh
End Sub

Private Sub sfrmPLMmaterialNum_Enter()
Me.sfrmPLMmaterialNum.Form.Requery
Me.sfrmPLMmaterialNum.Form.refresh
End Sub

Private Sub sfrmPLMmaterialSpecs_Enter()
Me.sfrmPLMmaterialSpecs.Form.Requery
Me.sfrmPLMmaterialSpecs.Form.refresh
End Sub

Private Sub sfrmPLMmaterialTypes_Enter()
Me.sfrmPLMmaterialTypes.Form.Requery
Me.sfrmPLMmaterialTypes.Form.refresh
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
