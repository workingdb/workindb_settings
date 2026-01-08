Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblTaskTracker", Me.Task_ID, Me.Title, "", "DELETED", "frmPLMsettings")

Dim db As Database
Set db = CurrentDb()
db.Execute "DELETE from tblTaskTracker where Task_ID = " & Me.Task_ID
Set db = Nothing
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub taskTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblTaskTracker", Me.Task_ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub values_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblTaskTracker", Me.Task_ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
