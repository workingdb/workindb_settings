Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cancelProject_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure? Cancelling a project cannot be undone.", vbYesNo, "Just making sure") = vbNo Then Exit Sub
If MsgBox("To cancel a project, you must have a file ready to upload. Do you have this ready?", vbYesNo, "Just making sure") = vbNo Then Exit Sub

Dim db As Database
Set db = CurrentDb

'kill approvals, open steps, open gates
db.Execute "delete * from tblPartTrackingApprovals where tableName = 'tblPartSteps' AND tableRecordId IN (SELECT recordId FROM tblPartSteps WHERE partProjectId = " & Me.recordId & ")"
db.Execute "delete * from tblPartSteps WHERE closeDate Is Null and status <> 'Closed' AND partProjectId = " & Me.recordId
DoEvents

'for gates that still have open steps (should be just one), put a completed date. Kill the other gates.
Dim rsGates As Recordset
Set rsGates = db.OpenRecordset("SELECT * FROM tblPartGates WHERE actualDate Is Null AND projectId = " & Me.recordId)

Do While Not rsGates.EOF
    If DCount("recordId", "tblPartSteps", "partGateId = " & rsGates!recordId) > 0 Then
        rsGates.Edit
        rsGates!actualDate = Date
        rsGates.Update
    Else
        rsGates.Delete
    End If
    rsGates.MoveNext
Loop

'notify team
Call addStatusChangeStep("Cancelled")
MsgBox "Please ensure Sales team member is on this email", vbInformation, "Double Check"

Call registerWdbUpdates("tblPartProject", Me.partNumber, "Part Project", Me.partNumber, "Cancelled", "frmPartProjectActions")
Call registerPartUpdates("tblPartProject", Me.recordId, "Project Status", Me.partProjectStatus, "Cancelled", Me.partNumber)
dbExecute "UPDATE tblPartProject SET projectStatus = 3 WHERE recordId = " & Me.recordId
Me.Requery

MsgBox "All done.", vbInformation, "It is finished."

exitThis:
rsGates.Close
Set rsGates = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub closeProject_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure? Force closing a project cannot be undone", vbYesNo, "Just making sure") = vbNo Then Exit Sub

'only allow if transfer ECO is implemented
Dim revID, errorTxt As String
errorTxt = ""

revID = idNAM(Me.partNumber, "NAM")
If revID = "" Then errorTxt = "Part Number Not Found, cannot close"

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT [CHANGE_NOTICE] from ENG_ENG_REVISED_ITEMS where [REVISED_ITEM_ID] = " & revID & _
    " AND [CANCELLATION_DATE] IS NULL AND [IMPLEMENTATION_DATE] IS NOT NULL AND [CHANGE_NOTICE] IN (SELECT [CHANGE_NOTICE] FROM ENG_ENG_ENGINEERING_CHANGES WHERE [CHANGE_ORDER_TYPE_ID] = 72)", dbOpenSnapshot)

If rs1.RecordCount = 0 Then errorTxt = "No Implemented Transfer ECO found, cannot close"

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Nope"
    GoTo exitThis
End If

'kill approvals, set all gates to closed
db.Execute "delete * from tblPartTrackingApprovals where tableName = 'tblPartSteps' AND tableRecordId IN (SELECT recordId FROM tblPartSteps WHERE partProjectId = " & Me.recordId & ")"
db.Execute "UPDATE tblPartGates SET actualDate = Date() WHERE actualDate Is Null AND projectId = " & Me.recordId

'close all open steps
Dim rsSteps As Recordset
Set rsSteps = db.OpenRecordset("SELECT * FROM tblPartSteps WHERE closeDate Is Null AND status <> 'Closed' AND partProjectId = " & Me.recordId)

Do While Not rsSteps.EOF
    rsSteps.Edit
    rsSteps!lastUpdatedDate = Now
    rsSteps!lastUpdatedBy = Environ("username")
    rsSteps!status = "Closed"
    rsSteps!closeDate = Now
    rsSteps.Update
    rsSteps.MoveNext
Loop

Call registerWdbUpdates("tblPartProject", Me.partNumber, "Part Project", Me.partNumber, "Closed", "frmPartProjectActions")
Call registerPartUpdates("tblPartProject", Me.recordId, "Project Status", Me.partProjectStatus, "Closed", Me.partNumber)
dbExecute "UPDATE tblPartProject SET projectStatus = 4 WHERE recordId = " & Me.recordId
Me.Requery

MsgBox "All done.", vbInformation, "It is finished."

exitThis:
On Error Resume Next
rs1.Close
Set rs1 = Nothing
rsSteps.Close
Set rsSteps = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub deleteProject_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure? Deleting a project cannot be undone.", vbYesNo, "Just making sure") = vbNo Then Exit Sub

Dim db As Database
Set db = CurrentDb

Call registerWdbUpdates("tblPartProject", Me.partNumber, "Part Project", Me.partNumber, "Deleted", "frmPartProjectActions")

db.Execute "delete * from tblPartTrackingApprovals where tableName = 'tblPartSteps' AND tableRecordId IN (SELECT recordId FROM tblPartSteps WHERE partProjectId = " & Me.recordId & ")"
db.Execute "UPDATE tblPartAttachmentsSP SET fileStatus='deleting' where partProjectId = " & Me.recordId
db.Execute "delete * from tblPartProject where recordId = " & Me.recordId

MsgBox "All done.", vbInformation, "It is finished."

Set db = Nothing

DoCmd.Close acForm, Me.name

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Me.inProgress.Enabled = False
Me.onHold.Enabled = False
Select Case Me.partProjectStatus
    Case "On Hold"
        Me.inProgress.Enabled = True
    Case "In Progress"
        Me.onHold.Enabled = True
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.Number)
End Sub

Private Sub inProgress_Click()
On Error GoTo Err_Handler

'potentially more added in the future

If MsgBox("Are you sure? This will re-enable all notifications, trackers, reports, etc.", vbYesNo, "Just making sure") = vbNo Then Exit Sub

Call registerPartUpdates("tblPartProject", Me.recordId, "Project Status", Me.partProjectStatus, "In Progress", Me.partNumber)
dbExecute "UPDATE tblPartProject SET projectStatus = 1 WHERE recordId = " & Me.recordId
Me.Requery

Call registerWdbUpdates("tblPartProjects", Me.partNumber, "Part Project", Me.partNumber, "Marked In Progress", "frmPartProjectActions")
MsgBox "All done.", vbInformation, "It is finished."

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub onHold_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure? This will disable all notifications, trackers, reports, etc.", vbYesNo, "Just making sure") = vbNo Then Exit Sub
If MsgBox("To hold a project, you must have a file ready to upload. Do you have this ready?", vbYesNo, "Just making sure") = vbNo Then Exit Sub

'add step with file attachment and note
Call addStatusChangeStep("Placed On Hold")
MsgBox "Please ensure Sales team member is on this email", vbInformation, "Double Check"

Call registerPartUpdates("tblPartProject", Me.recordId, "Project Status", Me.partProjectStatus, "On Hold", Me.partNumber)
dbExecute "UPDATE tblPartProject SET projectStatus = 2 WHERE recordId = " & Me.recordId
Me.Requery

Call registerWdbUpdates("tblPartProject", Me.partNumber, "Part Project", Me.partNumber, "Marked On Hold", "frmPartProjectActions")
MsgBox "All done.", vbInformation, "It is finished."

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Function addStatusChangeStep(action As String) As String
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

Dim gateId
Dim rsSteps As Recordset, docuType As Long
gateId = DLookup("recordId", "tblPartGates", "actualDate is null AND projectId = " & Me.recordId)

If Nz(gateId, "") = "" Then
    gateId = Nz(DMax("recordId", "tblPartGates", "projectId = " & Me.recordId))
End If

docuType = 36

Set rsSteps = db.OpenRecordset("tblPartSteps")
rsSteps.addNew
rsSteps!partNumber = Me.partNumber
rsSteps!partProjectId = Me.recordId
rsSteps!partGateId = gateId
rsSteps!stepType = "Project " & action
rsSteps!openedBy = Environ("username")
rsSteps!status = "Closed"
rsSteps!openDate = Now
rsSteps!lastUpdatedDate = Now
rsSteps!lastUpdatedBy = Environ("username")
rsSteps!closeDate = Now
rsSteps!documentType = docuType
rsSteps.Update

TempVars.Add "statusChangeStepId", db.OpenRecordset("SELECT @@identity")(0).Value

'---UPLOAD FILE---
Dim FSO As Object, attachName
Dim fileExt As String, currentLoc As String, fullPath As String, attchFullFileName As String, errorText As String
Set FSO = CreateObject("Scripting.FileSystemObject")

'---GRAB THE FILE---
Dim strFile
With Application.FileDialog(msoFileDialogOpen)
    .Title = "Choose a File"
    .AllowMultiSelect = False
    .Show
    On Error Resume Next
    strFile = .SelectedItems(1)
End With
On Error GoTo Err_Handler

If Nz(strFile) = "" Then
    MsgBox "Please select a document to upload...", vbCritical, "Hold up"
    Exit Function
End If

currentLoc = strFile
fileExt = FSO.GetExtensionName(currentLoc)

Dim rsDocumentType As Recordset
Set rsDocumentType = db.OpenRecordset("SELECT * from tblPartAttachmentStandards WHERE recordId = " & docuType)

attachName = rsDocumentType!FileName & "-" & DMax("ID", "tblPartAttachmentsSP") + 1
attchFullFileName = Replace(attachName, " ", "_") & "." & fileExt
        
Dim gateNum As Long
gateNum = CLng(Right(Left(DLookup("gateTitle", "tblPartGates", "recordId = " & gateId), 2), 1))

Dim rsPartAtt As DAO.Recordset
Dim rsPartAttChild As DAO.Recordset2
Set db = CurrentDb
Set rsPartAtt = db.OpenRecordset("tblPartAttachmentsSP", dbOpenDynaset)

rsPartAtt.addNew
rsPartAtt!fileStatus = "Created"

rsPartAtt.Update
rsPartAtt.MoveLast

rsPartAtt.Edit
Set rsPartAttChild = rsPartAtt.Fields("Attachments").Value

rsPartAttChild.addNew
Dim fld As DAO.Field2
Set fld = rsPartAttChild.Fields("FileData")
fld.LoadFromFile (currentLoc)
rsPartAttChild.Update

rsPartAtt!partNumber = Me.partNumber
rsPartAtt!partStepId = TempVars!statusChangeStepId
rsPartAtt!partProjectId = Me.recordId
rsPartAtt!documentType = docuType
rsPartAtt!uploadedBy = Environ("username")
rsPartAtt!uploadedDate = Now()
rsPartAtt!attachName = attachName
rsPartAtt!attachFullFileName = attchFullFileName
rsPartAtt!fileStatus = "Uploading"
rsPartAtt!gateNumber = gateNum
rsPartAtt!documentTypeName = rsDocumentType!documentType
rsPartAtt!businessArea = rsDocumentType!businessArea
rsPartAtt.Update

Call registerPartUpdates("tblPartAttachmentsSP", TempVars!statusChangeStepId, "Step Attachment", attachName, "Uploaded", Me.partNumber, "Project " & action, Me.recordId)

Dim attachLink As String
attachLink = "https://nifcoam.sharepoint.com/sites/NewModelEngineering/Part%20Info/" & attchFullFileName

'---SEND EMAIL---
Dim emailBody As String, subjectLine As String
subjectLine = "Notification: Part " & action
emailBody = generateHTML(subjectLine, Me.partNumber & " has been " & action, "Status Change Document", "No extra details...", "", "", attachLink)

Dim rs2 As Recordset, strTo As String
Set rs2 = db.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNumber & "'", dbOpenSnapshot)
strTo = ""

Do While Not rs2.EOF
    strTo = strTo & getEmail(rs2!person) & ";"
    rs2.MoveNext
Loop

strTo = Left(strTo, Len(strTo) - 1)

Dim SendItems As New clsOutlookCreateItem
Set SendItems = New clsOutlookCreateItem

SendItems.CreateMailItem sendTo:=strTo, _
                             subject:=subjectLine, _
                             htmlBody:=emailBody
    Set SendItems = Nothing

rs2.Close
Set rs2 = Nothing

'---CLEANUP---
On Error Resume Next
Set fld = Nothing
rsPartAttChild.Close: Set rsPartAttChild = Nothing
rsPartAtt.Close: Set rsPartAtt = Nothing
rsDocumentType.Close: Set rsDocumentType = Nothing
rsSteps.Close: Set rsSteps = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError(Me.name, "addStatusChangeStep", Err.DESCRIPTION, Err.Number)
End Function
