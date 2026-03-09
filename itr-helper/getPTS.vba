Sub getPTS()
    Dim basePath As String
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim outputRow As Long
    Dim tryPaths As Variant
    Dim i As Integer
    Dim found As Boolean
    Dim fileExt As String

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("ITRHelper")
    Set fso = CreateObject("Scripting.FileSystemObject")

    basePath = ws.Range("B2").Value

    ' Ensure base path ends with backslash
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"

    ' List of folders to try
    tryPaths = Array("02 Used data (cut-off from reference)", "01 Reference data", "02 Used data", "02 Extracted data")
    found = False

    ' Try each folder until one is found
    For i = LBound(tryPaths) To UBound(tryPaths)
        folderPath = basePath & tryPaths(i)
        If fso.FolderExists(folderPath) Then
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        ws.Range("D11").Value = "Folder not found"
        ws.Range("D12").Value = "No files found"
        Exit Sub
    End If

    ' Update D11 with resolved path
    ws.Range("D11").Value = folderPath

    Set folder = fso.GetFolder(folderPath)
    outputRow = 12
    ws.Range("D12:D10000").ClearContents

    ' List .txt, .csv, .pts files
    For Each file In folder.Files
        fileExt = LCase(Right(file.Name, Len(file.Name) - InStrRev(file.Name, ".") + 1))
        If fileExt = ".txt" Or fileExt = ".csv" Or fileExt = ".pts" Then
            ws.Cells(outputRow, 4).Value = file.Name
            outputRow = outputRow + 1
        End If
    Next file

    If outputRow = 12 Then
        ws.Range("D12").Value = "No files found"
    End If
    Exit Sub

ErrHandler:
    ws.Range("D12").Value = "Error occurred"
End Sub

