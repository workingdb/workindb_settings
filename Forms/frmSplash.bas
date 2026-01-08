Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call logClick("Form_Load", Me.Module.name)

If CommandBars("Ribbon").Height > 100 Then CommandBars.ExecuteMso "MinimizeRibbon"
DoCmd.ShowToolbar "Ribbon", acToolbarNo

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT * FROM tblPermissions WHERE user = '" & Environ("username") & "'")

If rs1.RecordCount = 0 Then 'if user doesn't exist, do NOT add them
    Application.Quit
    Exit Sub
End If

rs1.Close
Set rs1 = Nothing
Set db = Nothing

DoCmd.OpenForm "frmSettings"
DoCmd.Close acForm, "frmSplash"
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub
