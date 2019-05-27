Attribute VB_Name = "DrillDown"
Dim Operation As String
Dim OpCol As String
Dim OpFig As Double
Dim ConcCounter As Long

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
            Exit For
        Else
            DetermineCalcComponent = Left(txtArr(i + 2), Len(txtArr(i + 2)) - 1)
        End If
    End If
    i = i + 1
Next x

If i = 1 Then
    DetermineCalcComponent = Mid(txtArr(0), 2, Len(txtArr(0)) - 1)
End If

For x = 1 To 20
    If Right(DetermineCalcComponent, 1) = ")" Or Right(DetermineCalcComponent, 1) = "," Then
        DetermineCalcComponent = Left(DetermineCalcComponent, Len(DetermineCalcComponent) - 1)
    Else
        Exit For
    End If
Next x

DetermineCalcComponent = DetermineCalcComponent
    
Operation = Left(DetermineCalcComponent, InStr(DetermineCalcComponent, "(") - 1)


End Function

Function createFormula(txt As String, ws As Worksheet) As Boolean

Dim tmp2 As String
Dim st As Integer
Dim fn As Integer

createFormula = False

st = InStr(txt, "(") + 1
If Left(txt, 1) = "=" Then
    fn = Len(txt) - 5
Else
    fn = Len(txt) - 4
End If
txt = Mid(txt, st, fn)
createFormula = Evaluate(txt)

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

Arr = filterArr(DetermineCalcComponent(txt, ws), ws)

If UBound(Arr) - ConcCounter = 2 Then
    Call applyFilter1(Arr(0), Arr(1), Arr(2))
End If

If UBound(Arr) - ConcCounter = 4 Then
    Call applyFilter2(Arr(0), Arr(1), Arr(2), Arr(3), Arr(4))
End If

If UBound(Arr) - ConcCounter = 6 Then
    Call applyFilter3(Arr(0), Arr(1), Arr(2), Arr(3), Arr(4), Arr(5), Arr(6))
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
    txtArr(0) = Operation & "(" & txtArr(1)
End If
 '
i = 0

'set sheet and first range
    st = InStr(txtArr(i), "!") + 1
    fn = Len(txtArr(i)) + 1
    tmp = Mid(txtArr(i), st, fn - st)
    st = InStr(txtArr(i), "(") + 1
    fn = InStr(Mid(txtArr(i), st), "!") + st - 1
    txtArr(i) = Replace(Mid(txtArr(i), st, fn - st), "'", "")
    ReDim Preserve txtArr(UBound(txtArr))
    txtArr(i + 1) = tmp
'    txtArr(i) = Mid(txtArr(i), st, fn - st)
i = 1

conc = False

Dim counter As Long
ConcCounter = 0

For x = 1 To UBound(txtArr) - ConcCounter
    If (concperm = False And i Mod 2 = 1) Or InStr(txtArr(i), txtArr(0)) Then
        'get filter col as number (string format)
        st = InStr(txtArr(i), ":") + 1
        fn = Len(txtArr(i)) + 1
        txtArr(i - ConcCounter) = wColNum(txtArr(0), Mid(txtArr(i), st, fn - st))
    Else
        'get filter val as string
        If InStr(txtArr(i), "(") Then
            txtArr(i - ConcCounter) = txtArr(i) & "," & txtArr(i + 1)
            conc = True
            concperm = True
        End If
            If IsError(Evaluate(txtArr(i))) Then
                If IsError(Evaluate(txtArr(i) & ")")) Then
                    If IsError(Evaluate(Left(txtArr(i), Len(txtArr(i)) - 1))) Then
                        txtArr(i - ConcCounter) = Evaluate(Left(txtArr(i), Len(txtArr(i)) - 2))
                    Else
                        txtArr(i - ConcCounter) = (Evaluate(Left(txtArr(i), Len(txtArr(i)) - 1)))
                    End If
                Else
                    txtArr(i - ConcCounter) = Evaluate(txtArr(i) & ")")
                End If
            Else
                On Error GoTo IsText
                txtArr(i - ConcCounter) = Evaluate(txtArr(i))
            End If
        On Error GoTo 0
        GoTo Continue
IsText:
        st = InStr(txtArr(i), """") + 1
        fn = InStr(Mid(txtArr(i), st + 2), """") + st + 1
        txtArr(i - ConcCounter) = Mid(txtArr(i), st, fn - st)
        On Error GoTo 0
Continue:
    End If
If conc = True Then
    x = x + 1
    i = i + 2
    conc = False
    ConcCounter = ConcCounter + 1
Else
    i = i + 1
End If
Next x

filterArr = txtArr

Exit Function

End Function

Public Function wColNum(sh As String, ColNm As String) As Long

    wColNum = Range(ColNm & 1).Column
    
Dim ws As Worksheet

Set ws = ActiveWorkbook.Sheets(sh)
    
For Each col In ws.Range("A:A").Columns
    If col.EntireColumn.Hidden = True Then
        wColNum = wColNum - 1
    End If
Next col
    
    
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


Dim r As Long
If ws.Range("A1").Value = "" Then
    If ws.Range("A2").Value = "" Then
        r = 3
    Else
        r = 2
    End If
Else
    r = 1
End If


On Error Resume Next
ws.ShowAllData
On Error GoTo 0



If col = col2 Then
    With ws.Range("A" & r & ":" & LastColumn(sh, CStr(r)) & LastRow(sh, "A"))
        .AutoFilter Field:=col, Criteria1:=crit, Operator:=xlAnd, Criteria2:=crit2
    End With
Else
    With ws.Range("A" & r & ":" & LastColumn(sh, CStr(r)) & LastRow(sh, "A"))
        .AutoFilter Field:=col, Criteria1:=crit
        .AutoFilter Field:=col2, Criteria1:=crit2
    End With
End If


ws.Activate

End Sub

Sub applyFilter3(sh As String, col As String, crit As String, col2 As String, crit2 As String, col3 As String, crit3 As String)

Dim ws As Worksheet

Set ws = ActiveWorkbook.Sheets(sh)

Dim r As Long
If ws.Range("A1").Value = "" Then
    If ws.Range("A2").Value = "" Then
        r = 3
    Else
        r = 2
    End If
Else
    r = 1
End If


On Error Resume Next
ws.ShowAllData
On Error GoTo 0

If col = col2 Then  'if col & col2 are the same column
    With ws.Range("A" & r & ":" & LastColumn(sh, "1") & LastRow(sh, "A"))
        .AutoFilter Field:=col, Criteria1:=crit, Operator:=xlAnd, Criteria2:=crit2 'col & col2 filter line
        .AutoFilter Field:=col3, Criteria1:=crit3 'col3 filter line
    End With
Else
    If col2 = col3 Then 'if col2 & col3 are the same column
        With ws.Range("A" & r & ":" & LastColumn(sh, "1") & LastRow(sh, "A"))
            .AutoFilter Field:=col2, Criteria1:=crit2, Operator:=xlAnd, Criteria2:=crit3 'col2 & col3 filter line
            .AutoFilter Field:=col, Criteria1:=crit 'col filter line
        End With
    Else
        With ws.Range("A" & r & ":" & LastColumn(sh, "1") & LastRow(sh, "A")) 'if all cols are different
            .AutoFilter Field:=col, Criteria1:=crit 'col filter line
            .AutoFilter Field:=col2, Criteria1:=crit2 'col2 filter line
            .AutoFilter Field:=col3, Criteria1:=crit3 'col3 filter ilne
        End With
    End If
End If


ws.Activate

End Sub

