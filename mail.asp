<%
If request("go")<>"sent" Then response.End 
dim CLStr,msg,mailserver,username,password,receive
CLStr=Chr(13) & Chr(10)
'���ڴ��޸������Ϣ
mailserver="smtp.163.com" '�ʾַ�������ַ��smtp��������ַ��
username="18627130892@163.com" 'smtp��������֤��½����������Ϊ�����ʼ��ĵ�ַ,�����ʼ���email��ַ��
password="hzwl0769" 'smtp��������֤���� �������������룩
receive="service@js-pass.com" '���ܷ�����Ϣ��email��ַ�����������ʼ������䣩
'�޸Ľ���
Set msg = Server.CreateObject("JMail.Message")
msg.Charset = "gb2312"
msg.logging = true '�����ʼ���־
msg.silent=True'����������󣬷���False��True
'msg.ContentType = "text/html"'�ʼ��ĸ�ʽΪHTML��ʽ
msg.Priority = 1 '�ʼ��ȼ���1Ϊ�Ӽ���3Ϊ��ͨ��5Ϊ�ͼ�
msg.MailServerUserName = username
msg.MailServerPassword = password 
msg.From = username 
msg.FromName = username
msg.AddRecipient (receive)
msg.Subject = "��վ��������:"&Request.Form("subject")
msg.HTMLBody = "��վ����"&CLStr&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>��˾����:"&Request.Form("FaqTitle")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>��˾��ַ:"&Request.Form("xingbie")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>������:"&Request.Form("Content")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>��˾�绰:"&Request.Form("shenfenzheng")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>��˾����:"&Request.Form("dianhua")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>�ж��绰:"&Request.Form("weixin")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>�����ʼ�:"&Request.Form("jiezhongriqi")&CLStr
msg.HTMLBody = msg.HTMLBody&"<br>��ϸ����:<br>"
msg.HTMLBody = msg.HTMLBody&"<div style='font:9pt;background-color:#eeeeee'>"&Request.Form("yimiaozhongleixuanze")&CLStr
msg.HTMLBody = msg.HTMLBody&""
If msg.Send (mailserver) Then 
Response.Write(" <script language=javascript>alert('���ͳɹ�');location='/'</script>")
else
Response.Write(" <script language=javascript>alert('����ʧ�ܣ�����ϸ����ʼ��������������Ƿ���ȷ��') </script>")
End If 
msg.close
set msg = nothing
%>