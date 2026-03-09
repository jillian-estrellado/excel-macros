Sub Geotiff_Email()
    Dim ws As Worksheet
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim TemplatePath As String
    Dim reportDate As String
    Dim reportLink As String
    Set ws = ThisWorkbook.Sheets("Geotiff")

    ' Define template path
    TemplatePath = "D:\ESJI\MACROS&TOOLS\EmailTemplates\A2LZCO-2883 Updated UTD  Diff Chart Image.oft"

    ' Get values from Geotiff sheet
    With ThisWorkbook.Sheets("Geotiff")
        reportDate = Format(ws.Range("B3").Value, "dd-mmmm-yyyy")
        reportLink = ws.Range("B4").Value
    End With

    ' Create Outlook instance and load email template
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItemFromTemplate(TemplatePath)

    ' Replace placeholders in HTML body
    With MailItem
        .HTMLBody = Replace(.HTMLBody, "[DATE]", reportDate)
        .HTMLBody = Replace(.HTMLBody, "[HYPERLINK]", "<a href='" & reportLink & "'>" & reportLink & "</a>")
        .Display ' .Send or .Display if send change template to include signature
    End With

    ' Cleanup
    Set MailItem = Nothing
    Set OutlookApp = Nothing
End Sub



