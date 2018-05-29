
<%

'参数说明
'Subject : 邮件标题
'MailAddress : 发件服务器的地址,如smtp.163.com
'Email : 收件人邮件地址
'Sender : 发件人姓名
'Content : 邮件内容
'Fromer : 发件人的邮件地址

Sub SendAction(subject, email, sender, content) 
Set JMail = Server.CreateObject("JMail.Message") 
JMail.Charset = "utf-8" ' 邮件字符集，默认为"US-ASCII"
JMail.From = strMailUser ' 发送者地址
JMail.FromName = sender' 发送者姓名
JMail.Subject =subject
JMail.MailServerUserName = strMailUser' 身份验证的用户名
JMail.MailServerPassword = strMailPass ' 身份验证的密码
JMail.Priority = 3
JMail.AddRecipient(email)
JMail.Body = content
JMail.Send(strMailAddress)
End Sub

'调用此Sub的例子
Dim strSubject,strEmail,strMailAdress,strSender,strContent,strFromer
strSubject = "官网留言-by"&Request("xmx")&Request("xmm")
strContent = "姓名:" & Request("xmx") &Request("xmm") & VbCrLf & "公司:" & Request("company") & VbCrLf &  "社区:" & Request("sq") &  VbCrLf &  "门牌:" & Request("dw") &  VbCrLf &  "面积:" & Request("mj") & VbCrLf &  "月租金:" & Request("rent") & VbCrLf & "电话:" & Request("tel") & VbCrLf & "邮箱:" & Request("mail") & VbCrLf & "留言:" & vbcrlf & Request("msg")
strSender = Request("Name")
strEmail = "4659489@qq.com" '这是收信的地址,可以改为其它的邮箱
strMailAddress = "smtp.qq.com" '我司企业邮局地址，请使用 mail.您的域名
strMailUser = "4659489@qq.com" '我司企业邮局用户名
strMailPass = "rg5549287" '邮局用户密码

'Call SendAction (strSubject,strEmail,strSender,strContent)

%>
