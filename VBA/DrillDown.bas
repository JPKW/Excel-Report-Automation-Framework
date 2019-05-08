Attribute VB_Name = "DrillDown"
Dim Operation As String
Dim OpCol As String
Dim OpFig As Double



Function txt(rng As Range) As String

txt = rng.Formula

End Function

Function DetermineCalcComponent(txt As String, ws As Worksheet) As String

Dim txtArr() As String
Dim i As Integer

i = 0

txtArr = Split(txt, vbLf)

For Each x In txtArr
    If (Left(txtArr(i), 4)) = "=IF(" Or (Left(txtArr(i), 3)) = "IF(" Then
        If createFormula(txtArr(i), ws) = True Then
            DetermineCalcComponent = Left(txtArr(i + 1), Len(txtArr(i + 1)) - 1)
        Else
            DetermineCalcComponent = Left(txtArr(i + 2), Len(txtArr(i + 2)) - 1)
        End If
    End If
    i = i + 1
Next x

Operation = Left(DetermineCalcComponent, InStr(DetermineCalcComponent, "(") - 1)

End Function

Function createFormula(txt As String, ws As Worksheet) As Boolean

Dim tmp2 As String
Dim st As Integer
Dim fn As Integer

createFormula = False

st = InStr(txt, "(") + 1
fn = InStr(Mid(txt, st), "=") + st - 1

tmp = Mid(txt, st, fn - st)

st = InStr(Mid(txt, fn), "=") + fn
fn = InStr(Mid(txt, st + 2), """") + st + 2

tmp2 = Mid(txt, st, fn - st)

myval = """" & ws.Range(tmp).Value & """"

If tmp2 = myval Then createFormula = True

End Function

Sub FilterButton()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim b As Object
Dim ws As Worksheet
Dim rng As Range
Dim txt As String

    Set b = ActiveSheet.Buttons(Application.Caller)
        txt = b.TopLeftCell.Formula
        OpFig = b.TopLeftCell.Value
        Set ws = ActiveSheet

Dim Arr() As String

Arr = filterArr(DetermineCalcComponent(txt, ws), ThisWorkbook.Sheets("Dashboard"))

If UBound(Arr) = 2 Or UBound(Arr) = 3 Then
    Call applyFilter1(Arr(0), Arr(1), Arr(2))
End If

If UBound(Arr) = 4 Or UBound(Arr) = 5 Then
    Call applyFilter2(Arr(0), Arr(1), Arr(2), Arr(3), Arr(4))
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

If Not Operation = "COUNTIFS" Then
    MsgBox ("This figure (" & Round(OpFig, 2) & ") is the " & LCase(Left(Operation, InStr(Operation, "IF") - 1)) & " of column: " & OpCol & " on this sheet with this filter applied")
End If

End Sub

Function filterArr(txt As String, ws As Worksheet) As String()

Dim i As Integer
Dim st As Integer
Dim fn As Integer
Dim txtArr() As String
Dim tmp As String
Dim calc As String
Dim conc As Boolean

txtArr = Split(txt, ",")


If Operation = "COUNTIFS" Then
    i = UBound(txtArr) + 1
    ReDim Preserve txtArr(UBound(txtArr) + 1)
    For x = 0 To i
        If i = 0 Then Exit For
        txtArr(i) = txtArr(i - 1)
    i = i - 1
    Next x
Else
    OpCol = Right(txtArr(0), Len(txtArr(0)) - InStr(txtArr(0), ":"))
    ReDim Preserve txtArr(UBound(txtArr))
    i = UBound(txtArr) - 1
    For x = 0 To i
        If x = i Then Exit For
        If x = 0 Then
            txtArr(x) = Operation & "(" & txtArr(x + 1)
        Else
            If x > 1 Then txtArr(i) = txtArr(i - 1)
        End If
        i = i - 1
    Next x
End If
 '
i = 0

'set sheet and first range
    st = InStr(txtArr(i), "!") + 1
    fn = Len(txtArr(i)) + 1
    tmp = Mid(txtArr(i), st, fn - st)
    st = InStr(txtArr(i), "(") + 1
    fn = InStr(Mid(txtArr(i), st), "!") + st - 1
    txtArr(i) = Mid(txtArr(i), st, fn - st)
    ReDim Preserve txtArr(UBound(txtArr))
    txtArr(i + 1) = tmp
'    txtArr(i) = Mid(txtArr(i), st, fn - st)
i = 1

conc = False
concperm = False
Dim counter As Long
counter = 0

For x = 1 To UBound(txtArr) - counter
    If (concperm = False And i Mod 2 = 1) Or (concperm = True And i Mod 2 = 0) Then 'if true get filter col, else get filter val
        'get filter col as number (string format)
        st = InStr(txtArr(i), ":") + 1
        fn = Len(txtArr(i)) + 1
        txtArr(i - counter) = wColNum(Mid(txtArr(i), st, fn - st))
    Else
        'get filter val as string
        If InStr(txtArr(i), "(") Then
            txtArr(i - counter) = txtArr(i) & "," & txtArr(i + 1)
            conc = True
            concperm = True
        End If
            If IsError(Evaluate(txtArr(i))) Then
                If IsError(Evaluate(txtArr(i) & ")")) Then
                    txtArr(i - counter) = (Evaluate(Left(txtArr(i), Len(txtArr(i)) - 1)))
                Else
                    txtArr(i - counter) = Evaluate(txtArr(i) & ")")
                End If
            Else
                On Error GoTo IsText
                txtArr(i - counter) = Evaluate(txtArr(i))
            End If
        On Error GoTo 0
        GoTo Continue
IsText:
        st = InStr(txtArr(i), """") + 1
        fn = InStr(Mid(txtArr(i), st + 2), """") + st + 1
        txtArr(i - counter) = Mid(txtArr(i), st, fn - st)
        On Error GoTo 0
Continue:
    End If
If conc = True Then
    x = x + 1
    i = i + 2
    conc = False
    counter = counter + 1
Else
    i = i + 1
End If
Next x

filterArr = txtArr

Exit Function

End Function

Public Function wColNum(ColNm)
    wColNum = Range(ColNm & 1).Column
End Function


Sub applyFilter1(sh As String, col As String, crit As String)

Dim ws As Worksheet

Set ws = ActiveWorkbook.Sheets(sh)

On Error Resume Next
ws.ShowAllData
On Error GoTo 0

ws.Range("A1:" & LastColumn(sh, "1") & LastRow(sh, "A")).AutoFilter Field:=col, Criteria1:=crit, Operator:=xlAnd

ws.Activate

End Sub

Sub applyFilter2(sh As String, col As String, crit As String, col2 As String, crit2 As String)

Dim ws As Worksheet

Set ws = ActiveWorkbook.Sheets(sh)

On Error Resume Next
ws.ShowAllData
On Error GoTo 0

If col = col2 Then
    With ws.Range("A1:" & LastColumn(sh, "1") & LastRow(sh, "A"))
        .AutoFilter Field:=col, Criteria1:=crit, Operator:=xlAnd, Criteria2:=crit2
    End With
Else
    With ws.Range("A1:" & LastColumn(sh, "1") & LastRow(sh, "A"))
        .AutoFilter Field:=col, Criteria1:=crit
        .AutoFilter Field:=col2, Criteria1:=crit2
    End With
End If


ws.Activate

End Sub

