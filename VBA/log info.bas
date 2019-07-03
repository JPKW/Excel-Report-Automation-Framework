Sub LogInformation(LogMessage As String, LogFileName as String)

Dim FileNum As Integer

FileNum = FreeFile ' next file number

Open LogFileName For Append As #FileNum ' creates the file if it doesn't exist

Print #FileNum, LogMessage ' write information at the end of the text file

Close #FileNum ' close the file

End Sub

