Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)
Body = "Hi, Tommy <br><br>&nbsp &nbsp " _
& "This refer to the CO CR which scheduled to release on " _
& "<br>&nbsp &nbsp Those aren't registered in CAB. The self-assessment as attached for your review. " _
& "<br><br>&nbsp &nbsp Proper verification and health check plan have been well defined for this CO. Thanks ahead !" _ 
& "<br><br><br> Thanks & regards, " _
& "<br><b> Simon Gan</b>"

objMail.To= "XXX"
objMail.Cc= "yyy"
objMail.Subject = "Mail Subject"
objMail.HTMLBody = Body
objMail.Attachments.Add "C:\user\"
objMail.Display   
'objOutlook.Quit 關閉Outlook 
Set objMail = Nothing
Set objOutlook = Nothing