Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Caption_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblLinks", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmLinks", Me.LinkCaption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Department_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblLinks", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmLinks", Me.LinkCaption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

If validate = False Then Me.Undo

End Sub

Function validate()

validate = False

Dim errorMsg As String
errorMsg = ""

If IsNull(Me.LinkCaption) Or Me.LinkCaption = "" Then errorMsg = "Caption / Name"
If IsNull(Me.type) Or Me.type = "" Then errorMsg = "Type"
If IsNull(Me.Department) Or Me.Department = "" Then errorMsg = "Department"
If IsNull(Me.Link) Or Me.Link = "" Then errorMsg = "Link"

If errorMsg <> "" Then
    MsgBox "Please fill out " & errorMsg, vbInformation, "Please fix"
    Exit Function
End If

validate = True

End Function

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = 'frmLinks'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Link_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblLinks", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmLinks", Me.LinkCaption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub newLink_Click()
On Error GoTo Err_Handler

DoCmd.GoToRecord , , acNewRec

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub obsolete_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblLinks", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmLinks", Me.LinkCaption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Org_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblLinks", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmLinks", Me.LinkCaption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Restricted_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblLinks", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmLinks", Me.LinkCaption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub searchBox_Change()
On Error GoTo Err_Handler

Me.filter = "Caption LIKE '*" & Me.searchBox & "*' OR Link LIKE '*" & Me.searchBox & "*'"
Me.FilterOn = True

Me.searchBox.SelStart = Me.searchBox.SelLength

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler
If Me.showClosedToggle.Value = True Then
        Me.filter = "[Obsolete] = true"
    Else
        Me.filter = "[Obsolete] = false"
End If

Me.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Type_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblLinks", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmLinks", Me.LinkCaption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
