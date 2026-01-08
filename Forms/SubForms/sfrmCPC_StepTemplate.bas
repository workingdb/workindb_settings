Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If IsNull(Me.ID) Then
    MsgBox "This is an empty record.", vbInformation, "Can't do that"
    Exit Sub
End If

Dim db As Database
Set db = CurrentDb()

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Dim oldIndex
    oldIndex = Me.indexOrder
    If oldIndex < DMax("indexOrder", "tblCPC_Steps_Template", "projectTemplateId = " & Me.projectTemplateId) Then
        db.Execute "UPDATE tblCPC_Steps_Template SET indexOrder = indexOrder - 1 WHERE projectTemplateId = " & Me.projectTemplateId & " AND indexOrder > " & oldIndex
    End If
    
    Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, "DELETE", Me.stepName, "DELETED", "sfrmCPC_StepTemplate")
    db.Execute "DELETE FROM tblCPC_Approvals_Template WHERE stepTemplateId = " & Me.ID
    db.Execute "DELETE FROM tblCPC_Steps_Template WHERE ID = " & Me.ID
    Me.Requery
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub documentType_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "sfrmCPC_StepTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub duration_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "sfrmCPC_StepTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

Me.indexOrder = Nz(DMax("indexOrder", "tblCPC_Steps_Template", "projectTemplateId = " & Me.projectTemplateId) + 1, 1)
Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
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

If IsNull(Me.ID) Then Exit Sub

If Me.indexOrder = DMax("indexOrder", "tblCPC_Steps_Template", "projectTemplateId = " & Me.projectTemplateId) Then Exit Sub

Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex + 1

Dim db As Database
Set db = CurrentDb()
db.Execute "UPDATE tblCPC_Steps_Template SET indexOrder = " & oldIndex & " WHERE projectTemplateId = " & Me.projectTemplateId & " AND indexOrder = " & newIndex
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, "indexOrder", oldIndex, newIndex, "sfrmCPC_StepTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub moveUp_Click()
On Error GoTo Err_Handler

If IsNull(Me.ID) Then Exit Sub
If Me.indexOrder = 1 Then Exit Sub
Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex - 1

Dim db As Database
Set db = CurrentDb()
db.Execute "UPDATE tblCPC_Steps_Template SET indexOrder = " & oldIndex & " WHERE projectTemplateId = " & Me.projectTemplateId & " AND indexOrder = " & newIndex
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, "indexOrder", oldIndex, newIndex, "sfrmCPC_StepTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub responsible_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "sfrmCPC_StepTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub stepActionButton_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "sfrmCPC_StepTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub title_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblCPC_Steps_Template", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "sfrmCPC_StepTemplate")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
