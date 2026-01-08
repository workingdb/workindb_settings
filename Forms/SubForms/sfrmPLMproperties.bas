Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Data_Type_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMproperties", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Drawing_Parameter_Name_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMproperties", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Drawing_Text_Name_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMproperties", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Name_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMproperties", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Input_Disabled_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMproperties", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Input_Required_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMproperties", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Property_Name_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMproperties", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
