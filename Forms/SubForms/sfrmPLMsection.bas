Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Numbering_Table_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMsection", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Office_Code_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMsection", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Section_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMsection", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Selection_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMsection", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
