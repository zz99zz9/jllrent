<%if request("action")="send" then%>
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
JMail.Charset = "gb2312" ' 邮件字符集，默认为"US-ASCII"
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
  
  GetUrl="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")   
'  If   Request.ServerVariables("QUERY_STRING")<>""   Then   GetURL=GetUrl&"?"&   Request.ServerVariables("QUERY_STRING")
 
  GetUrl=replace(GetUrl,"contact.asp","")
  GetUrl=replace(GetUrl,"http://www.jllresidential.cn","")
strArr=split(GetUrl,"/")  

'验证信息是否重复
  



'调用此Sub的例子
Dim strSubject,strEmail,strMailAdress,strSender,strContent,strFromer
strSubject = "PARK SHORE 报名-"&Request("uname")
strContent = "Name:" & Request("uname") & VbCrLf & "Tel:" & Request("uphone") & VbCrLf & "City:" & Request("ucity") & VbCrLf & "From:" & GetUrl'strArr(3)
strSender = Request("Name")
strEmail = "slevin.wang@ap.jll.com" '这是收信的地址,可以改为其它的邮箱
strMailAddress = "smtp.qq.com" '我司企业邮局地址，请使用 mail.您的域名
strMailUser = "4659489@qq.com" '我司企业邮局用户名
strMailPass = "rg5549287" '邮局用户密码


%>

<%

if request.cookies("uphone")=Request("uphone") then
response.Redirect("index.html?3")
	elseif len(Request("uphone"))<>11 then
response.Redirect("index.html?2")
	else
	'Call SendAction (strSubject,strEmail,strSender,strContent)
		Call SendAction (strSubject,"4659489@qq.com",strSender,strContent)
		response.cookies("uphone")=Request("uphone")
		response.Redirect("index.html?1")
	%>


<%end if
end if
%>