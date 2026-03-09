Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A2")) Is Nothing Then
        Call ListPDFFiles
        Call getPTS
    End If
End Sub
