
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
strSubject = "����סլ��������Ϣ-By "&Request("uname")
strContent = "Name:" & Request("uname") & VbCrLf & "Tel:" & Request("uphone") & VbCrLf & "City:" & Request("ucity") & VbCrLf & "From:" & GetUrl'strArr(3)
strSender = Request("Name")
strEmail = "4659489@qq.com" '�������ŵĵ�ַ,���Ը�Ϊ���������� Project.Sales@ap.jll.com
strMailAddress = "smtp.exmail.qq.com" '��˾��ҵ�ʾֵ�ַ����ʹ�� mail.��������
strMailUser = "jll@hitpointcloud.com" '��˾��ҵ�ʾ��û���
strMailPass = "Hit12345" '�ʾ��û�����


%>

<%

'Call SendAction (strSubject,"jll@hitpointcloud.com",strSender,strContent)
	%>


