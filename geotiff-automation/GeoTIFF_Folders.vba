Sub GeoTIFF_Folders()
    Dim ws As Worksheet
    Dim folderPaths(1 To 4) As String
    Dim fso As Object
    Dim shellApp As Object
    Dim i As Integer
    Dim sourceFile1 As String, sourceFile2 As String
    Dim destFile1 As String, destFile2 As String

    Set ws = ThisWorkbook.Sheets("Geotiff")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shellApp = CreateObject("Shell.Application")

    ' Build folder paths from values in Geotiff sheet
    folderPaths(1) = "Z:\10 QINSy Data\09 GeoTIFF\UTD_Image\" & _
                     Trim(ws.Range("D1").Value) & "\" & Trim(ws.Range("E1").Value)
                     
    folderPaths(2) = "S:\Favorites\A2LZCO\03e ABS\Support activities\Charts\_UTD Image\" & _
                     Trim(ws.Range("D1").Value) & "\" & Trim(ws.Range("E1").Value) & "\"
                     
    folderPaths(3) = "Z:\10 QINSy Data\09 GeoTIFF\UTD_Image\" & _
                     Trim(ws.Range("D1").Value) & "\" & Trim(ws.Range("E1").Value) & "\Mean\"
                     
    folderPaths(4) = "Z:\10 QINSy Data\09 GeoTIFF\UTD_Image\" & _
                     Trim(ws.Range("D1").Value) & "\" & Trim(ws.Range("E1").Value) & "\Diff\"

    ' Write only folderPaths(2) to (4) starting at B4
    Dim rowOffset As Integer
    rowOffset = 4
    
    For i = 2 To 4
        ws.Range("B" & rowOffset).Value = folderPaths(i)
        rowOffset = rowOffset + 1
    Next i


    ' Ensure parent folder exists (folderPaths(1))
    If Dir(folderPaths(1), vbDirectory) = "" Then
        MkDirRecursive folderPaths(1)
    End If

    ' Create remaining folders (folderPaths 2 to 4)
    For i = 2 To 4
        If Dir(folderPaths(i), vbDirectory) = "" Then
            MkDirRecursive folderPaths(i)
        End If
    Next i

    ' Define source files
    sourceFile1 = "Z:\99 TEMP\ESJI\GEOTIFF\Color Scale.png"
    sourceFile2 = "Z:\99 TEMP\ESJI\GEOTIFF\Color Scale_DIFF.png"

    ' Copy Color Scale.png to Mean folder
    If fso.FileExists(sourceFile1) Then
        destFile1 = fso.BuildPath(folderPaths(3), "Color Scale.png")
        fso.CopyFile sourceFile1, destFile1, True
    Else
        Debug.Print "Source file 1 not found."
    End If

    ' Copy Color Scale_DIFF.png to Diff folder
    If fso.FileExists(sourceFile2) Then
        destFile2 = fso.BuildPath(folderPaths(4), "Color Scale_DIFF.png")
        fso.CopyFile sourceFile2, destFile2, True
    Else
        Debug.Print "Source file 2 not found."
    End If

    ' Open folders
    'If folderPaths(2) <> "" Then Shell "explorer.exe """ & folderPaths(2) & """", vbNormalFocus
    'If folderPaths(3) <> "" Then Shell "explorer.exe """ & folderPaths(3) & """", vbNormalFocus
    MsgBox "Done. Go to Qinsy.", vbInformation
End Sub

' Helper function to create folders recursively
Sub MkDirRecursive(ByVal fullPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(fullPath) Then
        fso.CreateFolder fullPath
    End If
End Sub


