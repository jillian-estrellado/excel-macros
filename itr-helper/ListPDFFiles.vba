Sub ListPDFFiles()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim fileList As Collection
    Dim fileItem As Variant
    Dim i As Long
    
    'Set the worksheet
    Set ws = ThisWorkbook.Sheets("ITRHelper")
    
    'Get the folder path from cell B2
    folderPath = ws.Range("B2").Value
    
    'Ensure the folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    'Check if folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        Exit Sub
    End If
    
    'Create a collection to hold all file paths
    Set fileList = New Collection
    
    'Get PDF files recursively
    Call GetPDFFilesRecursive(folderPath, fileList)
    
    'Clear previous list in column J
    ws.Range("J:J").ClearContents
    
    'Write PDF file names to column J
    i = 1
    For Each fileItem In fileList
        ws.Cells(i, "J").Value = Mid(fileItem, InStrRev(fileItem, "\") + 1)
        i = i + 1
    Next fileItem
    

End Sub


Private Sub GetPDFFilesRecursive(ByVal folderPath As String, ByRef fileList As Collection)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    'Add all PDF files in the current folder
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".pdf" Then
            fileList.Add file.path
        End If
    Next file
    
    'Recursively check subfolders
    For Each subFolder In folder.SubFolders
        GetPDFFilesRecursive subFolder.path, fileList
    Next subFolder
End Sub



