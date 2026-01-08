Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub businessArea_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartAttachmentStandards", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartTrackingSettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub documentType_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartAttachmentStandards", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartTrackingSettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub fileName_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPartAttachmentStandards", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPartTrackingSettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
