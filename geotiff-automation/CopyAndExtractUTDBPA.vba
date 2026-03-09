Sub CopyAndExtractUTDBPA()
    Dim ws As Worksheet
    Dim sourceFolder As String
    Dim destinationFolder As String
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim mostRecentFile As Object
    Dim mostRecentDate As Date
    Dim rarPath As String
    Dim extractCommand As String
    
    Set ws = ThisWorkbook.Sheets("Geotiff")

    ' Set folder paths
    sourceFolder = "SOURCELINK" & ws.Range("D1").Value & "\"
    destinationFolder = "DESTINATIONLINK"

    ' Ensure folders end with "\"
    If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
    If Right(destinationFolder, 1) <> "\" Then destinationFolder = destinationFolder & "\"

    ' Initialize file system object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(sourceFolder)

    mostRecentDate = 0

    ' Find most recent .rar file
    For Each file In folder.Files
        If LCase(fileSystem.GetExtensionName(file.Name)) = "rar" Then
            If file.DateLastModified > mostRecentDate Then
                Set mostRecentFile = file
                mostRecentDate = file.DateLastModified
            End If
        End If
    Next file

    If Not mostRecentFile Is Nothing Then
        ' Copy the most recent .rar file
        fileSystem.CopyFile _
            Source:=mostRecentFile.path, _
            Destination:=destinationFolder & mostRecentFile.Name, _
            OverWriteFiles:=True

        rarPath = destinationFolder & mostRecentFile.Name

        ' Build WinRAR extraction command
        extractCommand = """C:\Program Files\WinRAR\WinRAR.exe"" x -o+ """ & rarPath & """ """ & destinationFolder & """"

        ' Run the extraction command
        Shell extractCommand, vbHide

   
    Else
        MsgBox "No .rar files found in source folder.", vbExclamation
    End If

    ' Cleanup
    Set mostRecentFile = Nothing
    Set folder = Nothing
    Set fileSystem = Nothing
End Sub


