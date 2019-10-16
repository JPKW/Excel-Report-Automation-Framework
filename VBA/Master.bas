Option Explicit


Sub Generate()

Dim WB As Workbook
Dim WS As Worksheet
Dim F as integer

Set WB = ThisWorkbook
Set WS = thisWB.Sheets("MASTER")


Dim genWB As Workbook
Dim genRun As String

Dim lrun As String

lrun = Format(Now(), "DD/MM/YYYY")



For F = 50 To 2 step -1

    If Not WS.Cells(F, 6).Value = "Y" Then GoTo Skip

    Application.EnableEvents= False
    Application.DisplayAlerts= False
    Application.AskToUpdateLinks = False
    Set genWB = Workbooks.Open(WS.Cells(F, 3).Value)
    Application.EnableEvents= True
    Application.DisplayAlerts= True
    Application.AskToUpdateLinks = True
                    
    genRun = "'" & WS.Cells(F, 4).Value
    
    Application.Run (genRun)

    Application.DisplayAlerts = False
    genWB.Close
    Application.DisplayAlerts = True
    
    Set genWB = Nothing
    
    Application.Wait (Now + TimeValue("0:00:01"))

Skip:

Next F


End Sub



