<%if request("action")="send" then%>
<%

'����˵��
'Subject : �ʼ�����
'MailAddress : �����������ĵ�ַ,��smtp.163.com
'Email : �ռ����ʼ���ַ
'Sender : ����������
'Content : �ʼ�����
'Fromer : �����˵��ʼ���ַ

Sub SendAction(subject, email, sender, content) 
Set JMail = Server.CreateObject("JMail.Message") 
JMail.Charset = "gb2312" ' �ʼ��ַ�����Ĭ��Ϊ"US-ASCII"
JMail.From = strMailUser ' �����ߵ�ַ
JMail.FromName = sender' ����������
JMail.Subject =subject
JMail.MailServerUserName = strMailUser' �����֤���û���
JMail.MailServerPassword = strMailPass ' �����֤������
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

'��֤��Ϣ�Ƿ��ظ�
  



'���ô�Sub������
Dim strSubject,strEmail,strMailAdress,strSender,strContent,strFromer
strSubject = "PARK SHORE ����-"&Request("uname")
strContent = "Name:" & Request("uname") & VbCrLf & "Tel:" & Request("uphone") & VbCrLf & "City:" & Request("ucity") & VbCrLf & "From:" & GetUrl'strArr(3)
strSender = Request("Name")
strEmail = "slevin.wang@ap.jll.com" '�������ŵĵ�ַ,���Ը�Ϊ����������
strMailAddress = "smtp.qq.com" '��˾��ҵ�ʾֵ�ַ����ʹ�� mail.��������
strMailUser = "4659489@qq.com" '��˾��ҵ�ʾ��û���
strMailPass = "rg5549287" '�ʾ��û�����


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