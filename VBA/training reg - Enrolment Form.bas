Option Compare Database

'----------------------------------------------------------------------------------------------------------------------------
Private Sub Cmd_Enrol_Staff_Click()

Dim frm As Form
Dim ctl As Control
Dim varItm As Variant
Dim ctl2 As Control
Dim varItm2 As Variant
Set frm = Me.Form
Set ctl = frm!list_Courses
Set ctl2 = frm!list_Staff


'validate info, prompt user for "Are you sure??"

Dim n As Long '0   'for progress bar
Dim y As Long '100 'for progress bar
n = 0   'for progress bar
y = ctl.ItemsSelected.count * ctl2.ItemsSelected.count 'for progress bar

DoCmd.Hourglass True    'hourglass on
SysCmd acSysCmdInitMeter, "working...", y 'progress bar on

For Each varItm In ctl.ItemsSelected 'for each course selected
    For Each varItm2 In ctl2.ItemsSelected  'for each staff selected
        Call Add_Data(ctl.Column(0, varItm), ctl2.Column(0, varItm2)) 'add enrolled line
        n = n + 1   'for progress bar
        SysCmd acSysCmdUpdateMeter, n   'for progress bar
    Next varItm2
Next varItm

SysCmd acSysCmdRemoveMeter  'progress bar off
DoCmd.Hourglass False   'hourglass off

End Sub

'----------------------------------------------------------------------------------------------------------------------------

Private Sub Add_Data(Course As String, Staff As String)

    On Error GoTo Error_Handler
    Dim db                    As DAO.Database
    Dim rs                    As DAO.Recordset

    If Me.Dirty = True Then Me.Dirty = False 'Save any unsaved data

    Set db = CurrentDb
    Set rs = db.OpenRecordset("t_Enrolled")

    With rs
        .AddNew
        ![Sedgwick Email Address] = Staff
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
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: cmd_AddRec_Click" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
    
End Sub

Private Sub Form_Open(Cancel As Integer)

With List_Team_FUp
    .SetFocus
    For x = Abs(.ColumnHeads) To (.ListCount - 1)
            .Selected(x) = True
    Next x
End With

With List_Course_FUp
    .SetFocus
    For x = Abs(.ColumnHeads) To (.ListCount - 1)
            .Selected(x) = True
    Next x
End With

With List_Staff_FUp
    .SetFocus
    For x = Abs(.ColumnHeads) To (.ListCount - 1)
            .Selected(x) = True
    Next x
End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------
Private Sub List_Team_FUp_AfterUpdate()

Call FUpQuery

End Sub

'----------------------------------------------------------------------------------------------------------------------------
Private Sub List_Staff_FUp_AfterUpdate()

Call FUpQuery

End Sub

'----------------------------------------------------------------------------------------------------------------------------

Private Sub List_Course_FUp_AfterUpdate()

Call FUpQuery

End Sub

'----------------------------------------------------------------------------------------------------------------------------
Private Sub FUpQuery()

Dim TName As String
Dim CName As String
Dim SName As String

Dim ctl As Control
Dim varItm As Variant
Dim x As Integer

x = 0
Set ctl = Me.List_Team_FUp
For Each varItm In ctl.ItemsSelected 'for each team selected
    If x = 0 Then
        TName = "'" & ctl.Column(0, varItm) & "'"
    Else
        TName = TName & ",'" & ctl.Column(0, varItm) & "'"
    End If
    x = x + 1
Next varItm

x = 0
Set ctl = Me.List_Course_FUp
For Each varItm In ctl.ItemsSelected 'for each course selected
    If x = 0 Then
        CName = "'" & ctl.Column(0, varItm) & "'"
    Else
        CName = CName & ",'" & ctl.Column(0, varItm) & "'"
    End If
    x = x + 1
Next varItm

x = 0
Set ctl = Me.List_Staff_FUp
For Each varItm In ctl.ItemsSelected 'for each staff selected
    If x = 0 Then
        SName = "'" & ctl.Column(0, varItm) & "'"
    Else
        SName = SName & ",'" & ctl.Column(0, varItm) & "'"
    End If
    x = x + 1
Next varItm

Dim qBuilder As String

qBuilder = "WHERE [t_Enrolled++].[Next Follow Up Date] <= #" & Format(Now(), "MM/DD/YYYY") & "# AND [t_Enrolled++].[Follow Period Level] < 2"

If TName <> "" Then qBuilder = qBuilder & " AND [t_Enrolled++].[Team Name] In (" & TName & ")"

If CName <> "" Then qBuilder = qBuilder & " AND [t_Enrolled++].[Course Name] In (" & CName & ")"

If SName <> "" Then qBuilder = qBuilder & " AND [t_Enrolled++].[Sedgwick Email Address] In (" & SName & ")"

qBuilder = qBuilder & " ORDER BY [Sedgwick Email Address];"

'filter
Me.List_FUp.RowSource = "SELECT [ID], [Sedgwick Email Address], [Course Name], [Due Date], [Last Followed Up Date], [Next Follow Up With] " & _
"FROM [t_Enrolled++] " & qBuilder

' Refresh the list box
Me.List_FUp.Requery

End Sub

'----------------------------------------------------------------------------------------------------------------------------
'#########################################################################################################################
'#######################################################OLD###############################################################
'#########################################################################################################################
Private Sub FUp_Button_Click()

If MsgBox("Sending update emails & updating records for next follow up dates." & vbnewlinew & vbNewLine & "Are you sure you want to proceed?", vbYesNo + vbQuestion) = vbYes Then
Else
    Exit Sub
End If

'follow up selected in FUp_List
Dim frm As Form
Dim ctl As Control
Dim varItm As Variant
Dim ID As Integer
Dim EmailTo As String
Dim CourseName As String
Dim DueDate As String
Dim EmailCC As String
Dim EmailSubject As String
Dim EmailBody As String
Dim FName As String
Dim DMethod As String
Dim FUpCount As Integer

Set frm = Me.Form
Set ctl = frm!List_FUp

'validate info, prompt user for "Are you sure??"

Dim n As Integer '0   'for progress bar
Dim y As Integer '100 'for progress bar
n = 0   'for progress bar
y = ctl.ItemsSelected.count 'for progress bar

DoCmd.Hourglass True    'hourglass on
SysCmd acSysCmdInitMeter, "working...", y 'progress bar on

For Each varItm In ctl.ItemsSelected 'for each course selected
        'get data
        ID = ctl.Column(0, varItm)
        EmailTo = ctl.Column(1, varItm)
        FName = Functions.GetFieldInfo("t_Staff", EmailTo, "First Name")
        EmailCC = ctl.Column(5, varItm)
        DueDate = ctl.Column(3, varItm)
        CourseName = ctl.Column(2, varItm)
        EmailSubject = CourseName & " - Overdue"
        DMethod = Functions.GetFieldInfo("t_Courses", CourseName, "Delivery Method")
        FUpCount = Functions.GetFieldInfo("t_Enrolled", CStr(ID), "Follow Period Level")
        EmailBody = "Dear " & FName & "," & vbNewLine & vbNewLine & "Our records indicate that you were due to complete the " & _
            CourseName & " course on: " & DueDate & " - our records are not yet showing that this course has been successfully completed" & vbNewLine & vbNewLine
        
        EmailBody = EmailBody & Functions.GetFieldInfo("t_DeliveryMethods", DMethod, "Follow up " & FUpCount + 1 & " Template")
        
        If Functions.sendEmail(EmailBody, EmailSubject, EmailTo, EmailCC) = True Then
            Call Update_t_Enrolled_FUp(ID, FUpCount + 1)
        End If
        
        n = n + 1   'for progress bar
        SysCmd acSysCmdUpdateMeter, n   'for progress bar
Next varItm

SysCmd acSysCmdRemoveMeter  'progress bar off
DoCmd.Hourglass False   'hourglass off

Me.Refresh

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------
'#######################################################CURRENT!###############################################################

Private Sub FUp_Button_Click()

If MsgBox("Sending update emails & updating records for next follow up dates." & vbnewlinew & vbNewLine & "Are you sure you want to proceed?", vbYesNo + vbQuestion) = vbYes Then
Else
    Exit Sub
End If

'follow up selected in FUp_List
Dim frm As Form
Dim ctl As Control
Dim varItm As Variant
Dim ID As Integer
Dim EmailTo As String
Dim CourseName As String
Dim DueDate As String
Dim EmailCC As String
Dim EmailSubject As String
Dim EmailBody As String
Dim FName As String
Dim DMethod As String
Dim FUpCount As Integer

Set frm = Me.Form
Set ctl = frm!List_FUp

'validate info, prompt user for "Are you sure??"

Dim n As Integer '0   'for progress bar
Dim y As Integer '100 'for progress bar
n = 0   'for progress bar
y = ctl.ItemsSelected.count 'for progress bar

DoCmd.Hourglass True    'hourglass on
SysCmd acSysCmdInitMeter, "working...", y 'progress bar on

dim x as long
dim z as long
Dim f as long
dim FStart as long

For Each x in ctl.ItemsSelected.Count

FStart = x
z = 0
		'get data
        ID = ctl.Column(0, x)
        EmailTo = ctl.Column(1, x)
        FName = Functions.GetFieldInfo("t_Staff", EmailTo, "First Name")
        EmailCC = ctl.Column(5, x)
        DueDate = ctl.Column(3, x)
        CourseName = ctl.Column(2, x)
        EmailSubject = CourseName & " - Overdue"
        DMethod = Functions.GetFieldInfo("t_Courses", CourseName, "Delivery Method")
        FUpCount = Functions.GetFieldInfo("t_Enrolled", CStr(ID), "Follow Period Level")
        EmailBody = "Dear " & FName & "," & vbNewLine & vbNewLine & "Our records indicate that you were due to complete the " & _
            CourseName & " course on: " & DueDate & " - our records are not yet showing that this course has been successfully completed" & vbNewLine & vbNewLine
        
		EmailBody = EmailBody & Functions.GetFieldInfo("t_DeliveryMethods", DMethod, "Follow up " & FUpCount + 1 & " Template")
        
		While ctl.Column(1, x +1) = EmailTo AND ctl.Column(5, x +1) = EmailCC
			CourseName = ctl.Column(2, x + 1)
			DMethod = Functions.GetFieldInfo("t_Courses", CourseName, "Delivery Method")
			EmailSubject = CourseName & " + " & EmailSubject
			EmailBody = EmailBody & vbNewLine & vbNewLine & Functions.GetFieldInfo("t_DeliveryMethods", DMethod, "Follow up " & FUpCount + 1 & " Template")
				x = x + 1
				z = z + 1
		Wend

		If Functions.sendEmail(EmailBody, EmailSubject, EmailTo, EmailCC) = True Then
			For f = FStart to FStart + z
				ID = ctl.Column(0, F)
				FUpCount = Functions.GetFieldInfo("t_Enrolled", CStr(ID), "Follow Period Level")
				Call Update_t_Enrolled_FUp(ID, FUpCount + 1)
			next f
		End If
        
        n = n + 1   'for progress bar
        SysCmd acSysCmdUpdateMeter, n   'for progress bar

Next x

SysCmd acSysCmdRemoveMeter  'progress bar off
DoCmd.Hourglass False   'hourglass off

Me.Refresh

End Sub


'-------------------------------------------------------------------------------------------------------------------------------------

Private Sub Update_t_Enrolled_FUp(ID As Integer, FUpPeriod As Integer)

    On Error GoTo Error_Handler
    Dim db                    As DAO.Database
    Dim rs                    As DAO.Recordset

    If Me.Dirty = True Then Me.Dirty = False 'Save any unsaved data

    Set db = CurrentDb
    Set rs = db.OpenRecordset("t_Enrolled")

While rs.EOF = False
    If rs.Fields.item(0).value = ID Then
        rs.Edit
        rs![Follow Period Level] = FUpPeriod
        rs![Last Followed Up Date] = Format(Now(), "DD/MM/YYYY")
        rs.Update
        rs.MoveLast
    End If
    rs.MoveNext
Wend

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
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: cmd_AddRec_Click" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
    
End Sub

'------------------------------------------------------------------------------------------------

Function GetUserFolder() As String
   Dim fldr As Object
   Dim txtFileName As String

   ' FOLDER PICKER
   Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

   With fldr
      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select folder for Excel output."

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        txtFileName = .SelectedItems(1)
      Else
        MsgBox "No File Picked!", vbExclamation
        txtFileName = ""
      End If
   End With

   ' RETURN FOLDER NAME
   GetUserFolder = txtFileName
   
End Function

'------------------------------------------------------------------------------------------------------------------------------

Private Sub EnrolQuery()

Dim TName As String
Dim CName As String
Dim SName As String

Dim ctl As Control
Dim varItm As Variant
Dim x As Integer

x = 0

Set ctl = Me.enrol_ListTeam
For Each varItm In ctl.ItemsSelected 'for each team selected
    If x = 0 Then
        TName = "'" & ctl.Column(0, varItm) & "'"
    Else
        TName = TName & ",'" & ctl.Column(0, varItm) & "'"
    End If
    x = x + 1
Next varItm

x = 0
Set ctl = Me.optManFilter
For Each varItm In ctl.ItemsSelected 'for each course selected
    If x = 0 Then
        CName = "'" & ctl.Column(0, varItm) & "'"
    Else
        CName = CName & ",'" & ctl.Column(0, varItm) & "'"
    End If
    x = x + 1
Next varItm

Dim qBuilder As String

qBuilder = "WHERE [Sedgwick Email Address] <> 'Create New'"

If TName <> "" Then qBuilder = qBuilder & " AND [Team Name] In (" & TName & ")"

If CName <> "" Then qBuilder = qBuilder & " AND [Is Manager?] In (" & CName & ")"

qBuilder = qBuilder & " ORDER BY [Sedgwick Email Address];"

'filter
Me.List_FUp.RowSource = "SELECT [Sedgwick Email Address]" & _
"FROM t_Staff " & qBuilder

' Refresh the list box
Me.List_FUp.Requery

End Sub

'------------------------------------------------------------------------------------------------------------------------------

'Generate Executive Report
Private Sub Command45_Click()

Dim user_excel_fldr As String

    ' CALL FUNCTION
    user_excel_fldr = Functions.BrowseForFolder
    If user_excel_fldr = "" Then Exit Sub

    'SPECIFY ONE TABLE
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "t_ExecutiveReport", _
       user_excel_fldr & "\" & Format(Now(), "YYYYMMDD") & " Executive Report.xlsx", True


End Sub

