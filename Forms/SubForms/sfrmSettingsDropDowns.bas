Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Me.values.Locked = True

Dim filtVal As String
If userData("developer") = False Then
    filtVal = " AND dColumnDept = '" & userData("Dept") & "'"
End If

Me.selectColumn.RowSource = "SELECT dcolumns, dColumnDept FROM tblDropDownsSP WHERE dcolumns Is Not Null" & filtVal & " ORDER BY dcolumns;"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = '" & "frmSettingsDropdowns" & "'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub selectColumn_AfterUpdate()
On Error GoTo Err_Handler

Me.values.Locked = False
Me.values.ControlSource = Me.selectColumn.Value

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub values_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblDropDownsSP", Me.recordId, Me.values.ControlSource, Me.ActiveControl.OldValue, Me.ActiveControl, "frmSettingsDropdowns")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
