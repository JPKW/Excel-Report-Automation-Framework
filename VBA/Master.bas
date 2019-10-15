Option Explicit


Sub Generate()

Dim WB As Workbook
Dim WS As Worksheet

Set WB = ThisWorkbook
Set WS = thisWB.Sheets("MASTER")


Dim genWB As Workbook
Dim genRun As String

Dim lrun As String

lrun = Format(Now(), "DD/MM/YYYY")



For F = 50 To 2

If Not WS.Cells(F, 6).Value = "Y" Then Next F

    Set genWB = Workbooks.Open(WS.Cells(F, 3).Value)
    
    genRun = WS.Cells(F, 4).Value
    
    Application.Run (genWB.Name & "!" & genRun)

    Application.DisplayAlerts = False
    genWB.Close
    Application.DisplayAlerts = True
    
    Set genWB = Nothing
    
    Application.Wait (Now + TimeValue("0:00:01"))


Next F



End Sub



