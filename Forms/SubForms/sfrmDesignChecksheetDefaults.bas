Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub category_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmDesignChecksheetDefaults", Me.reviewItem)

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
    If oldIndex < DMax("indexOrder", "tblDesignChecksheetDefaults") Then
        db.Execute "UPDATE tblDesignChecksheetDefaults SET indexOrder = indexOrder - 1 WHERE indexOrder > " & oldIndex
    End If
    Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, "DELETE", Me.reviewItem, "DELETED", "frmDesignChecksheetDefaults", Me.reviewItem)
    db.Execute "DELETE FROM tblDesignChecksheetDefaults WHERE recordId = " & Me.recordId
    
    Me.Requery
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub designResponsible_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmDesignChecksheetDefaults", Me.reviewItem)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub drawingType_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmDesignChecksheetDefaults", Me.reviewItem)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

Me.indexOrder = Nz(DMax("indexOrder", "tblDesignChecksheetDefaults") + 1, 1)
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

If Me.indexOrder = DMax("indexOrder", "tblDesignChecksheetDefaults") Then Exit Sub

Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex + 1

Dim db As Database
Set db = CurrentDb()
db.Execute "UPDATE tblDesignChecksheetDefaults SET indexOrder = " & oldIndex & " WHERE indexOrder = " & newIndex
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, "indexOrder", oldIndex, newIndex, "frmDesignChecksheetDefaults", Me.reviewItem)

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
db.Execute "UPDATE tblDesignChecksheetDefaults SET indexOrder = " & oldIndex & " WHERE indexOrder = " & newIndex
Set db = Nothing
Me.indexOrder = newIndex
Me.Dirty = False
    
Me.Requery
Me.OrderBy = "indexOrder Asc"
Me.OrderByOn = True

Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, "indexOrder", oldIndex, newIndex, "frmDesignChecksheetDefaults", Me.reviewItem)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub partType_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmDesignChecksheetDefaults", Me.reviewItem)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub reviewItem_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblDesignChecksheetDefaults", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmDesignChecksheetDefaults", Me.reviewItem)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
