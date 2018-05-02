<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/xgconn.asp"-->
<%
BigClassName=request("BigClassName")
	  BigClassId=request("BigClassId")
	  SmallClassName=request("SmallClassName")
	  SmallClassId=request("SmallClassId")
	  Product_Id=request("Product_Id")
	  Title=request("Title")
	  uptime=request("uptime")
	  picalt=request("picalt")
	  Descriptio=request("Description")
	  Keywords=request("Keywords")
	  DefaultPicUrl=request("DefaultPicUrl")
	  DefaultPicUrl2=request("DefaultPicUrl2")
	  DefaultPicUrl3=request("DefaultPicUrl3")
      content=request("content")
	  Orderid=request("Orderid")
      Passed=request("Passed")
	  Elite=request("Elite")
	  action=trim(request("action"))
	  pro_id=request("pro_id")
	   gid=request("gid")
	  dl=request("dl")
	  if BigClassId="" then BigClassId=0
		if SmallClassId="" then SmallClassId=0
		if pro_id="" then pro_id=0
		if Orderid="" then Orderid=0
	'新增内容
	enTitle=request("enTitle")
	  jgzj=request("jgzj")
	  jgjj=request("jgjj")
	  if jgzj="" then jgzj=0
	  if jgjj="" then jgjj=0
	  mj1=request("mj1")
	  mj2=request("mj2")
	  if mj1="" then mj1=0
	  if mj2="" then mj2=0
	  cfjname=request("cfjname")
	  cfjid=request("cfjid")
	  clxname=request("clxname")
	  clxid=trim(request("clxid"))
	  crqname=request("crqname")
	  crqid=trim(request("crqid"))
	  clbname=request("clbname")
	  content1=request("content1")
	  content2=request("content2")
	  clbid=request("clbid")
'	  ctdname=request("ctdname")
	  ctdid=","&trim(request("ctdid"))&","
'二次新增内容
	wz=request("wz")
	syq=request("syq")
	jgsj=request("jgsj")
	qt=request("qt")
	hx=request("hx")
	ckfsid=request("ckfsid")
	ckfsname=request("ckfsname")
		
select case action

Case "add"
set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [Table_Product] ",conn,1,3

oRs.Addnew
oRs("BigClassName")=Trim(BigClassName)
oRs("BigClassId")=Trim(BigClassId)
oRs("SmallClassName")=Trim(SmallClassName)
oRs("SmallClassId")=Trim(SmallClassId)
oRs("Product_Id")=Trim(Product_Id)
oRs("Title")=Trim(Title)
oRs("dl")=Trim(dl)
oRs("gid")=Trim(gid)
oRs("picalt")=Trim(picalt)
oRs("Keywords")=Trim(Keywords)
oRs("Description")=Trim(Descriptio)
oRs("UpDatetime")=Trim(uptime)
oRs("DefaultPicUrl")=Trim(DefaultPicUrl)
oRs("DefaultPicUrl2")=Trim(DefaultPicUrl2)
oRs("DefaultPicUrl3")=Trim(DefaultPicUrl3)
oRs("content")=Trim(content)
oRs("Orderid")=Trim(Orderid)
oRs("Passed")=Trim(Passed)
oRs("Elite")=Trim(Elite)
'新增内容
oRs("enTitle")=trim(enTitle)
oRs("content1")=Trim(content1)
oRs("content2")=Trim(content2)
oRs("jgzj")=trim(jgzj)
oRs("jgjj")=trim(jgjj)
oRs("mj1")=trim(mj1)
oRs("mj2")=trim(mj2)
oRs("cfjname")=trim(cfjname)
oRs("cfjid")=trim(cfjid)
oRs("clxname")=trim(clxname)
oRs("clxid")=trim(clxid)
oRs("clbname")=trim(clbname)
oRs("clbid")=trim(clbid)
oRs("crqname")=trim(crqname)
oRs("crqid")=trim(crqid)
'oRs("ctdname")=trim(tdname)
oRs("ctdid")=trim(ctdid)
'二次新增内容
oRs("wz")=trim(wz) '位置
oRs("syq")=trim(syq)
oRs("jgsj")=trim(jgsj)
oRs("qt")=trim(qt)
oRs("hx")=trim(hx)
oRs("ckfsid")=trim(ckfsid)
oRs("ckfsname")=trim(ckfsname)
oRs("cg")=now()
oRs.Update

oRs.close
set oRs=Nothing
if Passed="false" then
response.write ("addcg")
else
response.write ("add")
end if
Case "mod"

set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [Table_Product] where ArticleId="&pro_id,conn,1,3


oRs("BigClassName")=Trim(BigClassName)
oRs("BigClassId")=Trim(BigClassId)
oRs("SmallClassName")=Trim(SmallClassName)
oRs("SmallClassId")=Trim(SmallClassId)
oRs("Product_Id")=Trim(Product_Id)
oRs("Title")=Trim(Title)
oRs("UpDatetime")=Trim(uptime)
oRs("DefaultPicUrl")=Trim(DefaultPicUrl)
oRs("DefaultPicUrl2")=Trim(DefaultPicUrl2)
oRs("DefaultPicUrl3")=Trim(DefaultPicUrl3)
oRs("content")=Trim(content)
oRs("Orderid")=Trim(Orderid)
oRs("Passed")=Trim(Passed)
oRs("Elite")=Trim(Elite)
oRs("gid")=Trim(gid)
oRs("picalt")=Trim(picalt)
oRs("dl")=Trim(dl)
'新增内容
oRs("Keywords")=Trim(Keywords)
oRs("Description")=Trim(Descriptio)
oRs("enTitle")=trim(enTitle)
oRs("content1")=Trim(content1)
oRs("content2")=Trim(content2)
oRs("jgzj")=trim(jgzj)
oRs("jgjj")=trim(jgjj)
oRs("mj1")=trim(mj1)
oRs("mj2")=trim(mj2)
oRs("cfjname")=trim(cfjname)
oRs("cfjid")=trim(cfjid)
oRs("crqname")=trim(crqname)
oRs("crqid")=trim(crqid)
oRs("clxname")=trim(clxname)
oRs("clxid")=trim(clxid)
oRs("clbname")=trim(clbname)
oRs("clbid")=trim(clbid)
'oRs("ctdname")=trim(tdname)
oRs("ctdid")=trim(ctdid)
'二次新增内容
oRs("wz")=trim(wz)
oRs("syq")=trim(syq)
oRs("jgsj")=trim(jgsj)
oRs("qt")=trim(qt)
oRs("hx")=trim(hx)
oRs("ckfsid")=trim(ckfsid)
oRs("ckfsname")=trim(ckfsname)
oRs("cg")=now()
oRs.Update

oRs.close
set oRs=Nothing
if Passed="false" then
response.write ("modcg")
else
response.write ("mod")
end if

Case "del"
conn.execute "delete from [Table_Product] where ArticleId="&Trim(pro_id)
response.write ("del")
end select
%>