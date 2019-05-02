Attribute VB_Name = "DrillDown"


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
        Set ws = ActiveSheet

Dim Arr() As String

Arr = filterArr(DetermineCalcComponent(txt, ws), ThisWorkbook.Sheets("Dashboard"))

If UBound(Arr) = 2 Then
    Call applyFilter1(Arr(0), Arr(1), Arr(2))
End If

If UBound(Arr) = 4 Then
    Call applyFilter2(Arr(0), Arr(1), Arr(2), Arr(3), Arr(4))
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Function filterArr(txt As String, ws As Worksheet) As String()

Dim i As Integer
Dim st As Integer
Dim fn As Integer
Dim txtArr() As String
Dim tmp As String
Dim calc As String

txtArr = Split(txt, ",")

i = UBound(txtArr) + 1

ReDim Preserve txtArr(UBound(txtArr) + 1)
For x = 0 To i
    If i = 0 Then Exit For
    txtArr(i) = txtArr(i - 1)
i = i - 1
Next x

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

For x = 1 To UBound(txtArr)
    If i Mod 2 = 1 Then
        st = InStr(txtArr(i), ":") + 1
        fn = Len(txtArr(i)) + 1
        txtArr(i) = wColNum(Mid(txtArr(i), st, fn - st))
    Else
        If InStr(txtArr(i), """") = 0 Then 'if this portion does not contain """ (=0 means not found)
            st = InStr(txtArr(i), ",") + 1
            fn = InStr(Mid(txtArr(i), st + 2), ")") + st + 1
            calc = Mid(txtArr(i), st, fn - st)
            txtArr(i) = ActiveSheet.Range(calc).Value
        Else
            st = InStr(txtArr(i), """") + 1
            fn = InStr(Mid(txtArr(i), st + 2), """") + st + 1
            txtArr(i) = Mid(txtArr(i), st, fn - st)
        End If
    End If
i = i + 1
Next x

filterArr = txtArr

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

With ws.Range("A1:" & LastColumn(sh, "1") & LastRow(sh, "A"))
    .AutoFilter Field:=col, Criteria1:=crit
    .AutoFilter Field:=col2, Criteria1:=crit2
End With

ws.Activate

End Sub


