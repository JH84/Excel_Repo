Attribute VB_Name = "EmailAll"

Sub Email_Distell()

Application.ScreenUpdating = False

    Dim oApp As Object, _
    oMail As Object, _
    FileName As String, MailSub As String, MailTxt As String, Mailto As String, MailCC As String, sAttachment As String
                    
    Worksheets("distell").Select
            
    For Each cell In Columns("g").Cells.SpecialCells(xlCellTypeVisible)
        If cell.Value Like "*@*" Then
            Mailto = Mailto & ";" & cell.Value
        End If
    Next
    
    'Loop through the rows for "To"
    For Each cell In Columns("i").Cells.SpecialCells(xlCellTypeVisible)
        If cell.Value Like "*@*" Then
            MailCC = MailCC & ";" & cell.Value
        End If
    Next
        
    MailSub = Range("c9").Value
    MailTxt = "<p style='font-family:calibri;font-size:15'>" & "Good day" & "<br><br>Please find attached the Daily Refrigeration Dashboard for, Store 18, Distell.<br><br>Regards" & "<br><img src='N:\DASHBOARDS_AND_REPORTS\TECHNICAL MONITORING\PnP C\Door Reports\Userform\Pmonitoring.jpg'" & "<p>"
    
    Application.ScreenUpdating = False
                    
     'Creates and shows the outlook mail item
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    With oMail
        .To = Mailto
        .cc = MailCC
        .Bcc = MailBCC
        .Subject = MailSub
        .Body = MailTxt
        .htmlbody = MailTxt
        .SentOnBehalfOfName = "PlantM@energypartners.co.za"
        On Error Resume Next
        .Attachments.Add Range("d1").Value
     '   .Attachments.Add Range("d2").Value
     '   .Attachments.Add Range("d3").Value
        .Display
    End With
    
    Application.ScreenUpdating = True
    Set oMail = Nothing
    Set oApp = Nothing
    
Worksheets("distell p1").Select

Application.ScreenUpdating = True

End Sub

