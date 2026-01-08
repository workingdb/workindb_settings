Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub checkItem_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartMeetingTemplates", Me.recordId, Me.checkItem, Me.ActiveControl.OldValue, Me.ActiveControl, Me.name, Me.cbMeetingType.column(1))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub cbMeetingType_AfterUpdate()
On Error GoTo Err_Handler

Me.filter = "meetingType = " & Me.cbMeetingType
Me.FilterOn = True

Me.meetingType.DefaultValue = Me.cbMeetingType

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

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
    If oldIndex < DMax("indexOrder", "tblPartMeetingTemplates") Then
        db.Execute "UPDATE tblPartMeetingTemplates SET indexOrder = indexOrder - 1 WHERE indexOrder > " & oldIndex & " AND meetingType = " & Me.meetingType
    End If
    Call registerWdbUpdates("tblPartMeetingTemplates", Me.recordId, "DELETE", Me.checkItem, "DELETED", Me.name, Me.cbMeetingType.column(1))
    db.Execute "DELETE FROM tblPartMeetingTemplates WHERE recordId = " & Me.recordId
    
    Me.Requery
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

Me.indexOrder = Nz(DMax("indexOrder", "tblPartMeetingTemplates", "meetingType = " & Me.cbMeetingType) + 1, 1)
Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
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

Private Sub moveDown_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If Me.indexOrder = DMax("indexOrder", "tblPartMeetingTemplates") Then Exit Sub

Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex + 1

Dim db As Database
Set db = CurrentDb()
db.Execute "UPDATE tblPartMeetingTemplates SET indexOrder = " & oldIndex & " WHERE indexOrder = " & newIndex & " AND meetingType = " & Me.meetingType
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblPartMeetingTemplates", Me.recordId, Me.checkItem & " indexOrder", oldIndex, newIndex, Me.name, Me.cbMeetingType.column(1))

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
db.Execute "UPDATE tblPartMeetingTemplates SET indexOrder = " & oldIndex & " WHERE indexOrder = " & newIndex & " AND meetingType = " & Me.meetingType
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblPartMeetingTemplates", Me.recordId, Me.checkItem & " indexOrder", oldIndex, newIndex, Me.name, Me.cbMeetingType.column(1))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
