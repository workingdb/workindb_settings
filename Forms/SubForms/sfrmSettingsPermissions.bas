Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Admin_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Beta_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub dataTag0_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Department_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub dept_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub designWOpermissions_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Edit_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub filterToMe_Click()
On Error GoTo Err_Handler

Me.filter = "user = '" & Environ("username") & "'"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub firstName_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Me.OrderBy = "ID Desc"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = 'frmPermissions'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Image106_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

Dim initials As String
initials = Left(Me.firstName, 1) & Left(Me.lastName, 1)

Call getAvatar(Me.User, initials)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub InActive_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub lastName_AfterUpdate()
On Error GoTo Err_Handler
Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub lblFirstName_Click()
On Error GoTo Err_Handler

Me.txtFirstName.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub lblLastName_Click()
On Error GoTo Err_Handler

Me.txtLastName.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub lblUser_Click()
On Error GoTo Err_Handler

Me.txtUser.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Level_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub newUser_Click()
On Error GoTo Err_Handler

Dim body As String, primaryMessage As String

Dim tblHeading As String, tblFooter As String, strHTMLBody As String


primaryMessage = "<a href = '\\data\mdbdata\WorkingDB\Batch\Shortcut\Working DB.lnk'>Click Here to Open WorkingDB</a>"

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & "WorkingDB Invitation" & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & "Once you open workingDB,<br/>please reply to this email so your permissions can be set.<br/>Please open inside of VMWare" & _
                                "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primaryMessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblFooter = "<table style=""width: 100%; margin: 0 auto; padding: 3em; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">Details</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;""><a href = 'https://nifcoam.sharepoint.com/:p:/r/sites/WorkingDB/Work Instructions/First Time User Overview.pptx?d=wde9823f7cab448a08f1e8f9ab8ba0a7e&csf=1&web=1&e=YmLuWG'>First Time User Overview</a></td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;""><a href = 'https://nifcoam.sharepoint.com/:p:/r/sites/WorkingDB/Work Instructions/Full Application.pptx?d=w2761e00c353245808c717e4d443f0884&csf=1&web=1&e=KV6VDq'>Full Application Work Instructions</a></td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & "The shortcut should auto-copy to your desktop" & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & "Sent by: " & getFullName & "</td></tr>" & _
                        "</tbody>" & _
                    "</table>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblFooter & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"


Call wdbEmail("", "brownj@us.nifco.com", "WorkingDB Invitation", strHTMLBody)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub removeFilter_Click()
On Error GoTo Err_Handler

Me.FilterOn = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub searchBox_Change()
On Error Resume Next

Me.filter = "firstName LIKE '*" & Me.searchBox & "*' OR lastName LIKE '*" & Me.searchBox & "*' OR user LIKE '*" & Me.searchBox & "*'"
Me.FilterOn = True

Me.searchBox.SelStart = Me.searchBox.SelLength

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler

If Me.showClosedToggle.Value = True Then
        Me.filter = "[Inactive] = true"
    Else
        Me.filter = "[Inactive] = false"
End If

Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Training_Mode_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub userEmail_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPermissions", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPermissions", Me.User)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
