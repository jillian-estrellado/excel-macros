Sub GetRQfolders()
    Dim folderPath As String
    Dim folderDialog As FileDialog
    Dim fso As Object, folder As Object
    Dim subFolder As Object
    Dim row As Long
    Dim existingPaths As Object
    Dim ws As Worksheet
    
 
    
    ' Only proceed if the active sheet is "RQFolders"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RQFolders")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'RQFolders' does not exist.", vbExclamation
        Exit Sub
    End If
    
    ' Create FileDialog to select folder
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If folderDialog.Show <> -1 Then Exit Sub
    folderPath = folderDialog.SelectedItems(1)
    
    ' Initialize FileSystemObject and Dictionary
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set existingPaths = CreateObject("Scripting.Dictionary")
    
    ' Read existing values in column A to skip duplicates
    row = 1
    Do While ws.Cells(row, 1).Value <> ""
        existingPaths(ws.Cells(row, 1).Value) = True
        row = row + 1
    Loop
    
    ' Start writing from the next empty row
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    ' Recursive function to scan folders
    Call ScanFolders(folder, existingPaths, row, ws)
    
    MsgBox "Done collecting ITR folders.", vbInformation
End Sub

Private Sub ScanFolders(ByVal folder As Object, ByRef existingPaths As Object, ByRef row As Long, ByRef ws As Worksheet)
    Dim subFolder As Object
    If InStr(1, folder.Name, "ITR", vbTextCompare) > 0 Then
        If Not existingPaths.Exists(folder.path) Then
            ws.Cells(row, 1).Value = folder.path
            existingPaths.Add folder.path, True
            row = row + 1
        End If
    End If
    
    For Each subFolder In folder.SubFolders
        ScanFolders subFolder, existingPaths, row, ws
    Next subFolder
End Sub

