<%
If request("go")<>"sent" Then response.End 
dim CLStr,msg,mailserver,username,password,receive
CLStr=Chr(13) & Chr(10)
'请在此修改相关信息
mailserver="smtp.163.com" '邮局服务器地址（smtp服务器地址）
username="18627130892@163.com" 'smtp服务器验证登陆名（用来做为代发邮件的地址,代发邮件的email地址）
password="hzwl0769" 'smtp服务器验证密码 （代发邮箱密码）
receive="service@js-pass.com" '接受反馈信息的email地址（用来接收邮件的信箱）
'修改结束
Set msg = Server.CreateObject("JMail.Message")
msg.Charset = "gb2312"
msg.logging = true '启用邮件日志
msg.silent=True'屏蔽例外错误，返回False或True
'msg.ContentType = "text/html"'邮件的格式为HTML格式
msg.Priority = 1 '邮件等级，1为加急，3为普通，5为低级
msg.MailServerUserName = username
msg.MailServerPassword = password 
msg.From = username 
msg.FromName = username
msg.AddRecipient (receive)
msg.Subject = "网站留言主题:"&Request.Form("subject")
msg.HTMLBody = "网站留言"&CLStr&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>公司名称:"&Request.Form("FaqTitle")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>公司地址:"&Request.Form("xingbie")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>联络人:"&Request.Form("Content")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>公司电话:"&Request.Form("shenfenzheng")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>公司传真:"&Request.Form("dianhua")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>行动电话:"&Request.Form("weixin")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>电子邮件:"&Request.Form("jiezhongriqi")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>详细内容:<br>"
msg.HTMLBody = msg.HTMLBody&"<div style='font:9pt;background-color:#eeeeee'>"&Request.Form("yimiaozhongleixuanze")&CLStr
msg.HTMLBody = msg.HTMLBody&""
If msg.Send (mailserver) Then 
Response.Write(" <script language=javascript>alert('发送成功');location='/'</script>")
else
Response.Write(" <script language=javascript>alert('发送失败，请仔细检查邮件服务器的设置是否正确！') </script>")
End If 
msg.close
set msg = nothing
%>