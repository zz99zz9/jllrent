<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/xgconn.asp"-->
<%
	  classid=request("classid")
	  cname=request("cname")
	  ename=request("ename")
	  position=request("position")
	  tel=request("tel")
	  summary=request("summary")
	  pic=request("DefaultPicUrl")
      mail=request("mail")
	  info=request("info")
	  oid=request("oid")
	uptime=trim(request("uptime"))
	action=trim(request("action"))
	id=trim(request("id"))

		if classid="" then classid=0

		if Oid="" then Oid=0

		
select case action

Case "add"
set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [team] ",conn,1,3

oRs.Addnew

oRs("cname")=trim(cname)
oRs("ename")=trim(ename)
oRs("position")=Trim(position)
oRs("pic")=Trim(pic)
oRs("summary")=Trim(summary)
oRs("oid")=Trim(oid)
oRs("uptime")=Trim(uptime)
oRs("tel")=Trim(tel)
oRs("info")=Trim(info)
oRs("cid")=Trim(classid)
oRs("mail")=Trim(mail)

oRs.Update

oRs.close
set oRs=Nothing
response.write ("add")

Case "mod"

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [team] where id="&id,conn,1,3
oRs("cname")=trim(cname)
oRs("ename")=trim(ename)
oRs("position")=Trim(position)
oRs("pic")=Trim(pic)
oRs("summary")=Trim(summary)
oRs("oid")=Trim(oid)
oRs("uptime")=Trim(uptime)
oRs("tel")=Trim(tel)
oRs("info")=Trim(info)
oRs("cid")=Trim(classid)
oRs("mail")=Trim(mail)
oRs.Update

oRs.close
set oRs=Nothing
response.write ("mod")
Case "del"
conn.execute "delete from [prod] where prod_Id="&Trim(pid)
response.write ("del")
end select
%>