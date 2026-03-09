Sub CompressCopyToCitrix()
    Dim rarPath As String
    Dim folder1 As String, folder2 As String
    Dim archive1 As String, archive2 As String
    Dim archivePath1 As String, archivePath2 As String
    Dim destinationFolder As String, extraFolder As String
    Dim fso As Object
    Dim wsh As Object
    Dim i As Integer
    Dim sourceFiles As Variant
    Dim filePath As String, fileName As String
    
    ' Initialize FileSystemObject and WScript.Shell
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    
    ' === Set WinRAR path ===
    rarPath = "C:\Program Files\WinRAR\rar.exe"
    
    ' === Get values from worksheet "GeoTIFF" ===
    With ThisWorkbook.Sheets("GeoTIFF")
        folder1 = Trim(.Range("B5").Value)
        folder2 = Trim(.Range("B6").Value)
        destinationFolder = Trim(.Range("B7").Value)
        extraFolder = Trim(.Range("B4").Value)
        archive1 = "2883_UTD (BPA)_" & .Range("E1").Value
        archive2 = "2883_UTD (BPA)_" & .Range("E1").Value & "_DIFF"
    End With
    
    ' === Validate inputs ===
    If folder1 = "" Or folder2 = "" Or destinationFolder = "" Then
        MsgBox "One or more required paths are empty (B4, B5, B6).", vbCritical
        Exit Sub
    End If
    
    ' === Ensure trailing backslash ===
    If Right(folder1, 1) <> "\" Then folder1 = folder1 & "\"
    If Right(folder2, 1) <> "\" Then folder2 = folder2 & "\"
    If Right(destinationFolder, 1) <> "\" Then destinationFolder = destinationFolder & "\"
    
    ' === Define full archive paths ===
    archivePath1 = folder1 & archive1 & ".rar"
    archivePath2 = folder2 & archive2 & ".rar"
    
    ' === Run WinRAR with wait ===
    wsh.Run """" & rarPath & """ a -ep1 -r -y """ & archivePath1 & """ """ & folder1 & "*.*""", 0, True
    wsh.Run """" & rarPath & """ a -ep1 -r -y """ & archivePath2 & """ """ & folder2 & "*.*""", 0, True
    
    ' === Copy archives to destination ===
    sourceFiles = Array(archivePath1, archivePath2)
    
    On Error Resume Next
    For i = LBound(sourceFiles) To UBound(sourceFiles)
        filePath = sourceFiles(i)
        If fso.FileExists(filePath) Then
            fileName = fso.GetFileName(filePath)
            fso.CopyFile filePath, destinationFolder & fileName, True
        Else
            MsgBox "File not found: " & filePath, vbExclamation
        End If
    Next i
    On Error GoTo 0
    
    ' === Open folders B4 and B7 ===
    If fso.FolderExists(extraFolder) Then
        wsh.Run "explorer.exe """ & extraFolder & """", 1, False
    End If
    If fso.FolderExists(destinationFolder) Then
        wsh.Run "explorer.exe """ & destinationFolder & """", 1, False
    End If
    
End Sub

