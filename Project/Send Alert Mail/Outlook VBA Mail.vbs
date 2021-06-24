' 先添加Outlook资料库(VBA)

Sub open
'定义工作簿事件,打开自动或者点击
Application.Ontime TimeValue("16:59:59"),"SendMail"
End Sub

Sub SendMail()
Dim objOutlook As Outlook.Application
Dim objMail As MailItem

Set objOutlook = New Outlook.Application '创建objOutlook为Outlook应用程序对象
Set objMail = objOutlook.CreateItem(olMailItem)  '创建objMail为一个邮件对象


With objMail
  .To = Recipient        '收件人
  .CC = Recipientcc      '抄送
  .Subject = Subject     '标题
  .Body = Body           '正文
  .Attachments.Add File  '附件
  .Display               '展示
  .Send                  '发送
End With

Set objMail = Nothing
Set objOutlook = Nothing

End Sub

## 循环多个邮件一起发送 ##

Sub Send_Email_Macro()

Dim Info_Arr As Variant
Dim n%
Info_Arr = Sheet1.UsedRange
n = UBound(Info_Arr)

For i = 2 To n
  Call Send_Email_By_Outlook(i)
Next i

End Sub


Sub Send_Email_By_Outlook(ByVal i As Integer)

Recipient = Info_Arr(i, 1)              'A列第二行开始
Recipientcc = Info_Arr(i, 2)            'B列第二行开始
Subj = Info_Arr(i, 3)                   'C列第二行开始
Body = Info_Arr(i, 4)                   'D列第二行开始
File = Info_Arr(i, 5)                   'E列第二行开始
If Len(Dir(File)) = 0 Then File = ""    '判断文件是否存在,要写绝对路径


Dim objOutlook As Outlook.Application
Dim objMail As MailItem

Set objOutlook = New Outlook.Application
Set objMail = objOutlook.CreateItem(olMailItem)

With objMail
  .To = Recipient
  .CC = Recipientcc
  .Subject = Subj
  .Body = Body
  .Attachments.Add File
  .Send
End With

Set objMail = Nothing
Set objOutlook = Nothing

End Sub

## 以表格形式写入正文,调用公共函数 ##

With objMail
            .To = Cells(rng.Row, 5).Value '//收件人
            .Subject = "工资明细" '//主题
            .BodyFormat = 2
            .HTMLBody = RangetoHTML(sht2.Range("a1:d2"))
            .display
            .send
End With


Public Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
    With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=TempFile, _
        Sheet:=TempWB.Sheets(1).Name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
    "align=left x:publishsource=")
    TempWB.Close savechanges:=False
    Kill TempFile
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function