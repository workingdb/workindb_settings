Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addNew_Click()
On Error GoTo Err_Handler

Me.Material_Type_ID.SetFocus
Me.Form.Recordset.addNew
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Material_Spec_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMdropDownsMaterialSpec", Me.Material_Spec_ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Material_Type_ID_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMdropDownsMaterialSpec", Me.Material_Spec_ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
