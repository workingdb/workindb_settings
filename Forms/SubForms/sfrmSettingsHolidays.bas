Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addNew_Click()
On Error GoTo Err_Handler

Me.holidayName.SetFocus
Me.Form.Recordset.addNew

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Me.searchBox = Year(Date)
Me.filter = "year(holidayDate) = " & Me.searchBox
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = 'frmHolidays'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub holidayDate_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblHolidays", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmHolidays")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub holidayName_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblHolidays", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmHolidays")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub
If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Call registerWdbUpdates("tblHolidays", Me.recordId, "DELETED", Me.holidayName, "DELETED")
dbExecute "DELETE FROM tblHolidays WHERE [recordId] = " & Me.recordId

Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub srch_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
Me.filter = "year(holidayDate) = " & Me.searchBox
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
