Sub FillMedicationChartsFromExcel()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wdTable As Object
    Dim ws As Worksheet

    Dim folderPath As String
    Dim fileName As String
    Dim clientID As String

    Dim lastRow As Long
    Dim r As Long
    Dim wordRow As Long
    Dim clearRow As Long
    Dim c As Long

    Set ws = ThisWorkbook.Sheets("Medications")

folderPath = "C:\SampleProject\Merged Docs\"

    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder not found:" & vbCrLf & folderPath
        Exit Sub
    End If

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    wdApp.DisplayAlerts = 0

    fileName = Dir(folderPath & "C*.docx")

    Do While fileName <> ""

        clientID = Split(fileName, "_")(0)

        On Error Resume Next
        Set wdDoc = wdApp.Documents.Open( _
            FileName:=folderPath & fileName, _
            ReadOnly:=False, _
            AddToRecentFiles:=False _
        )

        If Err.Number <> 0 Then
            MsgBox "Could not open: " & fileName & vbCrLf & Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        Set wdTable = FindMedicationTable(wdDoc)

        If wdTable Is Nothing Then
            MsgBox "Medication table not found in " & fileName
            wdDoc.Close False
            GoTo NextFile
        End If

        For clearRow = 4 To 16
            For c = 1 To 10
                On Error Resume Next
                wdTable.Cell(clearRow, c).Range.Text = ""
                On Error GoTo 0
            Next c
        Next clearRow

        wordRow = 4
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        For r = 2 To lastRow

            If Trim(ws.Cells(r, 1).Value) = clientID Then

                If wordRow > wdTable.Rows.Count Then
                    wdTable.Rows.Add
                End If

                wdTable.Cell(wordRow, 1).Range.Text = ws.Cells(r, 4).Value
                wdTable.Cell(wordRow, 2).Range.Text = ws.Cells(r, 5).Value
                wdTable.Cell(wordRow, 3).Range.Text = ws.Cells(r, 6).Value
                wdTable.Cell(wordRow, 4).Range.Text = ws.Cells(r, 7).Value
                wdTable.Cell(wordRow, 5).Range.Text = ws.Cells(r, 8).Value
                wdTable.Cell(wordRow, 6).Range.Text = ws.Cells(r, 9).Value
                wdTable.Cell(wordRow, 7).Range.Text = ws.Cells(r, 10).Value
                wdTable.Cell(wordRow, 8).Range.Text = ws.Cells(r, 11).Value
                wdTable.Cell(wordRow, 9).Range.Text = ws.Cells(r, 12).Value
                wdTable.Cell(wordRow, 10).Range.Text = ws.Cells(r, 13).Value

                wordRow = wordRow + 1

            End If

        Next r

        wdDoc.Save
        wdDoc.Close False

NextFile:
        Set wdDoc = Nothing
        fileName = Dir

    Loop

    wdApp.Quit
    Set wdApp = Nothing

    MsgBox "Medication charts filled."

End Sub


Function FindMedicationTable(wdDoc As Object) As Object

    Dim tbl As Object

    For Each tbl In wdDoc.Tables

        If InStr(1, tbl.Range.Text, "MY MEDICATIONS", vbTextCompare) > 0 _
        Or (InStr(1, tbl.Range.Text, "Medication", vbTextCompare) > 0 _
        And InStr(1, tbl.Range.Text, "Prescribed By", vbTextCompare) > 0) Then

            Set FindMedicationTable = tbl
            Exit Function

        End If

    Next tbl

    Set FindMedicationTable = Nothing

End Function
