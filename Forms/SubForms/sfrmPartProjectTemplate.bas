Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then
    MsgBox "This is an empty record.", vbInformation, "Can't do that"
    Exit Sub
End If

Dim db As Database
Set db = CurrentDb()

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Dim oldIndex
    oldIndex = Me.indexOrder
    If oldIndex < DMax("indexOrder", "tblPartStepTemplate", "gateTemplateId = " & Me.gateTemplateId) Then
        db.Execute "UPDATE tblPartStepTemplate SET indexOrder = indexOrder - 1 WHERE gateTemplateId = " & Me.gateTemplateId & " AND indexOrder > " & oldIndex
    End If
    
    Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, "DELETE", Me.Title, "DELETED", "frmPartProjectTemplate")
    db.Execute "DELETE FROM tblPartStepTemplateApprovals WHERE stepTemplateId = " & Me.recordId
    db.Execute "DELETE FROM tblPartStepTemplate WHERE recordId = " & Me.recordId
    Me.Requery
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub documentType_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub duration_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

Me.indexOrder = Nz(DMax("indexOrder", "tblPartStepTemplate", "gateTemplateId = " & Me.gateTemplateId) + 1, 1)
Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.txtCF = Me.recordId

Me.sfrmPartTrackingTemplateApprovals.Form.filter = "[stepTemplateId] = " & Me.recordId
Me.sfrmPartTrackingTemplateApprovals.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub moveDown_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If Me.indexOrder = DMax("indexOrder", "tblPartStepTemplate", "gateTemplateId = " & Me.gateTemplateId) Then Exit Sub

Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex + 1

Dim db As Database
Set db = CurrentDb()
db.Execute "UPDATE tblPartStepTemplate SET indexOrder = " & oldIndex & " WHERE gateTemplateId = " & Me.gateTemplateId & " AND indexOrder = " & newIndex
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, "indexOrder", oldIndex, newIndex, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub moveUp_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub
If Me.indexOrder = 1 Then Exit Sub
Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex - 1

Dim db As Database
Set db = CurrentDb()
db.Execute "UPDATE tblPartStepTemplate SET indexOrder = " & oldIndex & " WHERE gateTemplateId = " & Me.gateTemplateId & " AND indexOrder = " & newIndex
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, "indexOrder", oldIndex, newIndex, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub responsible_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub stepActionButton_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub title_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartStepTemplate", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartProjectTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
