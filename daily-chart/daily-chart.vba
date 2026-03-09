Sub DailyChart_Email()
    Dim ws As Worksheet
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim TemplatePath As String
    Dim reportDate As String
    Dim reportLink As String
    Dim Date1 As String
    Dim Date2 As String
       
    Set ws = ThisWorkbook.Sheets("DailyChart")

    ' Define template path
    TemplatePath = "D:\TEMPLATE.oft"

    ' Get values from Daily Chart sheet
    With ws
        Date1 = Format(.Range("B3").Value, "yymmdd")
        Date2 = Format(.Range("B3").Value, "dd mmmm yyyy")
    End With

    ' Create Outlook instance and load email template
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItemFromTemplate(TemplatePath)

    ' Replace placeholders in HTML body and update subject
    With MailItem
        .HTMLBody = Replace(.HTMLBody, "[DATE1]", Date1)
        .HTMLBody = Replace(.HTMLBody, "[DATE2]", Date2)
        .Subject = "SUBJECT HERE " & Date1  ' Update subject line
        .Display  ' Use .Send if you want it to send automatically
    End With

    ' Cleanup
    Set MailItem = Nothing
    Set OutlookApp = Nothing
End Sub

