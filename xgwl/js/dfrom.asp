<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>


<%
db="../../%#Data$%Ba#se%%/%#Data$%Ba#se%%.mdb"
connstr="Provider=Microsoft.Jet.OLEDB.4.0;data source="&Server.MapPath(db)
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.Open connstr
'类修改
dname=request("dname")
dtel=request("dtel")
dmail=request("dmail")
if dtel<>"" then
  conn.execute"INSERT INTO [lyb] (lyname,lyTel,lydate,lyemail) VALUES ('"&dname&"','"&dtel&"','"&now()&"','"&dmail&"')"
  end if

%>
