Attribute VB_Name = "Functions"

'#######################################################################################
'################## Created by Joerg Wood (github.com/PushyFantastic) ##################
'#######################################################################################

'############################################################################################################################
'gives last column in specified row
Function LastColumn(wsheet As String, row As String) As String

Dim ws As Worksheet
Set ws = ActiveWorkbook.Sheets(wsheet)

LastColumn = Split(Columns(ws.Cells(row, Columns.Count).End(xlToLeft).Column).Address(, False), ":")(1)

End Function

'############################################################################################################################
'gives last row in specified column
Function LastRow(wsheet As String, col As String) As Long
Dim ws As Worksheet
Set ws = ActiveWorkbook.Sheets(wsheet)

LastRow = ws.Cells(Rows.Count, col).End(xlUp).row

End Function

'############################################################################################################################
'for caluclating work hours duration between two DateTimes
Function WorkHours(StartDateTime, EndDateTime, Bus_Hrs_Start As Date, Bus_Hrs_End As Date, Holidays)
Dim d1 As Date, d2 As Date, wf As WorksheetFunction
Dim t1 As Date, t2 As Date
Dim HrsElapsed As Double

Set wf = Application.WorksheetFunction

d1 = DateValue(StartDateTime)
d2 = DateValue(EndDateTime)
t1 = TimeValue(StartDateTime)
t2 = TimeValue(EndDateTime)

    If t1 < TimeValue(Bus_Hrs_Start) Then
        t1 = TimeValue(Bus_Hrs_Start)
    End If
    If t1 > TimeValue(Bus_Hrs_End) Then
        t1 = TimeValue(Bus_Hrs_End)
    End If
    If t2 < TimeValue(Bus_Hrs_Start) Then
        t2 = TimeValue(Bus_Hrs_Start)
    End If
    If t2 > TimeValue(Bus_Hrs_End) Then
        t2 = TimeValue(Bus_Hrs_End)
    End If

If Weekday(d1) = 7 Or Weekday(d1) = 1 Or wf.CountIfs(Holidays, d1) > 0 Then '7 = sat, 1 = sun, countifs = public holiday
    d1 = wf.WorkDay(d1, 1, Holidays)
    t1 = TimeValue(Bus_Hrs_Start)
End If

If Weekday(d2) = 7 Or Weekday(d2) = 1 Or wf.CountIfs(Holidays, d2) > 0 Then '7 = sat, 1 = sun, countifs = public holiday
    d2 = wf.WorkDay(d2, 1, Holidays)
    t2 = TimeValue(Bus_Hrs_Start)
End If
  
If wf.NetworkDays(d1, d2, Holidays) > 1 Then
    HrsElapsed = ((wf.NetworkDays(d1, d2, Holidays) - 2) * 8) + TimeDiff(t1, Bus_Hrs_End) + TimeDiff(Bus_Hrs_Start, t2)
Else
    HrsElapsed = TimeDiff(t1, t2)
End If

WorkHours = HrsElapsed

End Function

Function TimeDiff(StartTime As Date, EndTime As Date) 'this function is used in the above function

    TimeDiff = Abs(EndTime - StartTime) * 24
    
End Function

'############################################################################################################################
'this is useful for summing the absolute values of a range (eg for weighted scoring)
Function SumAbs(rng As Range) As Long

result = 0

    On Error Resume Next
    For Each element In rng
        result = result + Abs(element)
    Next element
    On Error GoTo 0

SumAbs = result

End Function

'############################################################################################################################
'This function will insert the date as "YYYYMMDD - " into the file name of a file path

Function InjectDate(FilePath As String) As String

Dim FPArray() As String

FPArray = Split(FilePath, "\")

For x = 0 To UBound(FPArray)

    If x = UBound(FPArray) Then
        InjectDate = InjectDate + Format(Now(), "YYYYMMDD") + " - " + FPArray(x)
    Else
        InjectDate = InjectDate + FPArray(x) + "\"
    End If

Next x

End Function


'############################################################################################################################
'############################################################################################################################
'############################################################################################################################

Sub ListBoxUpdate()

Dim fl As Range

If ActiveSheet.Shapes(Application.Caller).ControlFormat.ListFillRange = "AdjusterList" Then
    Set fl = ActiveWorkbook.Sheets("Validation").Range("FilterList")
Else
    Set fl = ActiveWorkbook.Sheets("Validation").Range("FilterList2")
End If



Dim flVal As String

flVal = ""

    Dim i As Long
    With ActiveSheet.Shapes(ActiveSheet.Shapes(Application.Caller).Name).OLEFormat.Object
        For i = 1 To .ListCount
            If .Selected(i) Then
                flVal = .List(i) & "|" & flVal 'item i selected
            End If
        Next i
    End With

fl = flVal


End Sub

