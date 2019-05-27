'team management.bas

Option Compare Database

Private Sub Cmd_Enrol_Staff_Click()

Dim frm As Form
Dim ctl As Control
Dim varItm As Variant
Dim ctl2 As Control
Dim varItm2 As Variant

Set frm = Forms!Enrolment
Set ctl = frm!list_Courses
Set ctl2 = frm!list_Staff

'validate info, prompt user for "Are you sure??"

For Each varItm In ctl.ItemsSelected 'for each course selected
        For Each varItm2 In ctl2.ItemsSelected  'for each staff selected
                Call Add_Data(ctl.Column(intI, varItm), ctl2.Column(intI2, varItm2)) 'add enrolled line
        Next varItm2
Next varItm

End Sub

Private Sub Add_Data(Course As String, staff As String)

    On Error GoTo Error_Handler
    Dim db                    As DAO.Database
    Dim rs                    As DAO.Recordset

    If Me.Dirty = True Then Me.Dirty = False 'Save any unsaved data

    Set db = CurrentDb
    Set rs = db.OpenRecordset("t_Enrolled")

    With rs
        .AddNew
        ![Sedgwick Email Address] = staff
        ![Course Name] = Course
        ![Enrolled Date] = Format(Now(), "DD/MM/YYYY")
        .Update
    End With

Error_Handler_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
    Exit Sub

Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: cmd_AddRec_Click" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
    
End Sub

'-------------------------------------------------------------------------------------------------------

'functions.bas
Option Compare Database

Function GetFieldInfo(T_Name As String, Primary_Key As String, FieldName As String) As String

Dim objRecordset As Object
Set objRecordset = CreateObject("ADODB.Recordset")
Dim i As Integer
Dim value As Variant

objRecordset.ActiveConnection = CurrentProject.Connection
objRecordset.Open (T_Name)

'loop through table fields
For i = 0 To objRecordset.Fields.Count - 1
    If FieldName = objRecordset.Fields.Item(i).Name Then Exit For
Next i



'find the target record
While objRecordset.EOF = False
'check for match
If objRecordset.Fields.Item(0).value = Primary_Key Then
'get value
    GetFieldInfo = objRecordset.Fields(i).value
    'exit loop
    objRecordset.MoveLast
End If
objRecordset.MoveNext
Wend


End Function

