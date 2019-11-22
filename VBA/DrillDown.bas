'#######################################################################################
'####################### Created by Joerg Wood (github.com/JPKW) #######################
'#######################################################################################



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

If i = 1 And Left(txtArr(0), 8) <> "=IFERROR" Then
    DetermineCalcComponent = Mid(txtArr(0), 2, Len(txtArr(0)) - 1)
End If

For x = 1 To 20
    If Right(DetermineCalcComponent, 1) = ")" Or Right(DetermineCalcComponent, 1) = "," Then
        DetermineCalcComponent = Left(DetermineCalcComponent, Len(DetermineCalcComponent))
    Else
        Exit For
    End If
Next x

If DetermineCalcComponent = "" Then
    If Left(txtArr(0), 8) = "=IFERROR" Then
        DetermineCalcComponent = Mid(txtArr(1), 1, Len(txtArr(1)) - 1)
    End If
End If

  
  
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

Call applyFilter(Arr())


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
Dim newArr() As String


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
                                ReDim Preserve newArr(0)
                                newArr(0) = Replace(Mid(txtArr(i), st, fn - st), "'", "")
    txtArr(i) = Replace(Mid(txtArr(i), st, fn - st), "'", "")
    ReDim Preserve txtArr(UBound(txtArr))
    txtArr(i + 1) = tmp
'    txtArr(i) = Mid(txtArr(i), st, fn - st)
i = 1

conc = False
Dim counter As Integer
Dim concZ As Integer
ConcCounter = 0


For x = 1 To UBound(txtArr) - ConcCounter
concZ = 0
    If (concperm = False And i Mod 2 = 1) Or InStr(txtArr(i), txtArr(0)) Then
        'get filter col as number (string format)
        st = InStr(txtArr(i), ":") + 1
        fn = Len(txtArr(i)) + 1
                                ReDim Preserve newArr(UBound(newArr) + 1)
                                newArr(UBound(newArr)) = wColNum(txtArr(0), Mid(txtArr(i), st, fn - st))
        txtArr(i - ConcCounter) = wColNum(txtArr(0), Mid(txtArr(i), st, fn - st))
    Else
        'get filter val as string
        If InStr(txtArr(i), "(") Then
            txtArr(i - ConcCounter) = txtArr(i) & "," & txtArr(i + 1)
            conc = True
            concperm = True
                If InStr(txtArr(i), ")") = 0 Then
                    txtArr(i) = txtArr(i - ConcCounter) & "," & txtArr(i + 2)
                    conc = True
                    concperm = True
                    concZ = 1
                End If
        End If
            If IsError(Evaluate(txtArr(i))) Then
                If IsError(Evaluate(txtArr(i) & ")")) Then
                    If IsError(Evaluate(Left(txtArr(i), Len(txtArr(i)) - 1))) Then
                                ReDim Preserve newArr(UBound(newArr) + 1)
                                newArr(UBound(newArr)) = Evaluate(Left(txtArr(i), Len(txtArr(i)) - 2))
                                txtArr(i - ConcCounter) = Evaluate(Left(txtArr(i), Len(txtArr(i)) - 2))
                    Else
                                ReDim Preserve newArr(UBound(newArr) + 1)
                                newArr(UBound(newArr)) = (Evaluate(Left(txtArr(i), Len(txtArr(i)) - 1)))
                        txtArr(i - ConcCounter) = (Evaluate(Left(txtArr(i), Len(txtArr(i)) - 1)))
                    End If
                Else
                                ReDim Preserve newArr(UBound(newArr) + 1)
                                newArr(UBound(newArr)) = Evaluate(txtArr(i) & ")")
                    txtArr(i - ConcCounter) = Evaluate(txtArr(i) & ")")
                End If
            Else
                On Error GoTo IsText
                                ReDim Preserve newArr(UBound(newArr) + 1)
                                newArr(UBound(newArr)) = Evaluate(txtArr(i))
                txtArr(i - ConcCounter) = Evaluate(txtArr(i))
            End If
        On Error GoTo 0
        GoTo Continue
IsText:
        st = InStr(txtArr(i), """") + 1
        fn = InStr(Mid(txtArr(i), st + 2), """") + st + 1
                                ReDim Preserve newArr(UBound(newArr) + 1)
                                newArr(UBound(newArr)) = Mid(txtArr(i), st, fn - st)
        txtArr(i - ConcCounter) = Mid(txtArr(i), st, fn - st)
Continue:
On Error GoTo 0
    End If
    
    If conc = True Then
        x = x + 1 + concZ
        i = i + 2 + concZ
        conc = False
        ConcCounter = ConcCounter + 1 + concZ
    Else
        i = i + 1
    End If

Next x

filterArr = newArr

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

Sub applyFilter(filterArr() As String)

Dim ws As Worksheet

Set ws = ActiveWorkbook.Sheets(filterArr(0))

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

With ws.Range("A" & r & ":" & LastColumn(ws.Name, CStr(r)) & LastRow(ws.Name, "A"))
    For x = 1 To UBound(filterArr) - 1 Step 2
        If x = UBound(filterArr) - 1 Then 'if it's the last one
                    .AutoFilter Field:=filterArr(x), Criteria1:=filterArr(x + 1)
        Else
            If filterArr(x) = filterArr(x + 2) Then
                        .AutoFilter Field:=filterArr(x), Criteria1:=filterArr(x + 1), Operator:=xlAnd, Criteria2:=filterArr(x + 3) 'col & col2 filter line
                x = x + 2
            Else
                    .AutoFilter Field:=filterArr(x), Criteria1:=filterArr(x + 1)
            End If
        End If
    Next x
End With

ws.Activate

End Sub

