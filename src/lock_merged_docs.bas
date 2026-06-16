Sub FullyLockMergedDocs()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim folderPath As String
    Dim fileName As String
    Dim pw As String

pw = "SamplePassword123!"

folderPath = "C:\SampleProject\Merged Docs\"

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    wdApp.DisplayAlerts = 0

    fileName = Dir(folderPath & "*.docx")

    Do While fileName <> ""

        On Error Resume Next
        Set wdDoc = wdApp.Documents.Open( _
            FileName:=folderPath & fileName, _
            ReadOnly:=False, _
            AddToRecentFiles:=False)

        If Err.Number <> 0 Then
            MsgBox "Skipped: " & fileName & vbCrLf & Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        If wdDoc.ProtectionType <> -1 Then
            wdDoc.Unprotect Password:=pw
        End If

        wdDoc.Protect Type:=3, NoReset:=True, Password:=pw

        wdDoc.Save
        wdDoc.Close False

NextFile:
        Set wdDoc = Nothing
        fileName = Dir

    Loop

    wdApp.Quit

    MsgBox "All available merged docs locked with password."

End Sub
