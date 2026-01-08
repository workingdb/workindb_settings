Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addNew_Click()
On Error GoTo Err_Handler

Me.Material_Type.SetFocus
Me.Form.Recordset.addNew
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Private Sub Material_Type_AfterUpdate()
On Error GoTo Err_Handler

Call registerWdbUpdates("tblPLMdropDownsMaterialType", Me.Material_Type_ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "frmPLMsettings")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub

Public Sub Material_Type_Click()
On Error GoTo Err_Handler

If Me.Material_Type = "" Or IsNull(Me.Material_Type) Then
    Exit Sub
End If
If Form_frmPLMsettings.matFilter = True Then
    Dim matType As String
    matType = DLookup("[Material_Type_ID]", "tblPLMdropDownsMaterialType", "[Material_Type] = '" & Me.Material_Type & "'")
    Form_sfrmPLMmaterialGrades.filter = "[Material_Type_ID] = " & matType
    Form_sfrmPLMmaterialGrades.FilterOn = True
    Form_sfrmPLMmaterialSpecs.filter = "[Material_Type_ID] = " & matType
    Form_sfrmPLMmaterialSpecs.FilterOn = True
    
    Dim db As Database
    Dim rs1 As Recordset, rs2 As Recordset
    Set db = CurrentDb()
    Set rs2 = db.OpenRecordset("tblPLMdropDownsMaterialGrade", dbOpenSnapshot)
    rs2.filter = "[Material_Type_ID] = " & matType
    Set rs1 = rs2.OpenRecordset
    Dim matGrade() As Integer, lngCnt As Long: lngCnt = 0
    ReDim Preserve matGrade(0)
    Do While Not rs1.EOF
        ReDim Preserve matGrade(lngCnt)
        matGrade(lngCnt) = rs1![Material_Grade_ID]
        lngCnt = lngCnt + 1
        rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    Dim matGradeFilt As String
    If IsNull(matGrade(0)) Then
        matGradeFilt = "[Material_Grade_ID] = "
        GoTo skipFor
    End If
    
    matGradeFilt = "([Material_Grade_ID] = " & matGrade(0) & ")"
    
    If UBound(matGrade()) = 0 Then
        GoTo skipFor
    End If
    
    Dim i As Integer
    For i = 1 To UBound(matGrade())
        matGradeFilt = matGradeFilt & " OR ([Material_Grade_ID] = " & matGrade(i) & ")"
    Next
skipFor:
    Form_sfrmPLMmaterialNum.filter = matGradeFilt
    Form_sfrmPLMmaterialNum.FilterOn = True
Else
    Form_sfrmPLMmaterialGrades.FilterOn = False
    Form_sfrmPLMmaterialSpecs.FilterOn = False
    Form_sfrmPLMmaterialNum.FilterOn = False
End If
Set db = Nothing
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.Number)
End Sub
