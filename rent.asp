﻿<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%
dim tdkid
tdkid=2
%>
<!--#include file="./inc/tdk.asp"-->
<%dim bc,sc,fj,lx,lb,order,page
key=request.QueryString("key")
bc=request.QueryString("bc")
sc=request.QueryString("sc")
fj=request.QueryString("fj")
lx=request.QueryString("lx")
lb=request.QueryString("lb")
order=request.QueryString("order")
page=request.QueryString("page")
if bc="" then bc=0
if sc="" then sc=0
if fj="" then fj=0
if lx="" then lx=0
if lb="" then lb=0
if page="" then page=1
if order="" then order=0

%>
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
        <style>
        .ninfo .ntxt{border-top:0px;}
    </style>
<!--广告部份-->
<div class="navwz container container2">
<a href="http://www.jllresidential.cn">JLL</a> &gt; <a href="rent.asp">租赁</a><a href="jllmap.asp" class='hidden-xs hidden'><img src="xgwl/img/temp/dtbtn.gif" class="dtbtn b_h"></a>
</div>
<!--广告部份-->
<div class="cla container container2"> <%set rs=Server.CreateObject("ADODB.Recordset")
                           rs.Open "select * from [Table_ProBigClass] order by OrderId desc",conn,1,1%>
<span class="clat">城市:</span><span class="TAB_CLICKa"><a href="?<%call seaurl(0,0,fj,lx,lb,order,1)%>"  <%call ison(bc,0)%>>不限</a>

 <%do while not rs.eof%>
  <a href="?<%call seaurl(rs("BigClassId"),0,fj,lx,lb,order,1)%>" <%call ison(bc,rs("BigClassId"))%>><%=rs("BigClassName")%></a>
  <%rs.movenext

    loop
	rs.close
	set rs=nothing%>

</span><div class="c line1px"></div>
  <%if bc<>"" and bc<>0 then%>
    <%set rs=Server.CreateObject("ADODB.Recordset")
    rs.Open "select * from [Table_ProSmallClass] where BigClassID="&bc&" order by OrderId desc",conn,1,1%>
<span class="clat">区域:</span><span class="TAB_CLICKa"><a href="?<%call seaurl(0,0,fj,lx,lb,order,1)%>"  <%call ison(sc,0)%>>不限</a>

<%do while not rs.eof%>
  <a href="?<%call seaurl(bc,rs("SmallClassId"),fj,lx,lb,order,1)%>" <%call ison(sc,rs("SmallClassId"))%>><%=rs("SmallClassName")%></a>
  <%rs.movenext

    loop
	rs.close
	set rs=nothing%>
</span><div class="c line1px"></div>
<%end if%>
<%set rs=Server.CreateObject("ADODB.Recordset")
  rs.Open "select * from [Class_fj] order by oId desc,cid desc",conn,1,1%>
<span class="clat">租金:</span><span class="TAB_CLICKa"><a href="?<%call seaurl(bc,sc,0,lx,lb,order,1)%>" <%call ison(fj,0)%>>不限</a>

<%do while not rs.eof%>
     <a href="?<%call seaurl(bc,sc,rs("cid"),lx,lb,order,1)%>" rel="nofollow" <%call ison(fj,rs("cid"))%>><%=rs("cname")%></a>
       <%rs.movenext

    loop
	rs.close
	set rs=nothing%>

</span><div class="c line1px"></div>
<%set rs=Server.CreateObject("ADODB.Recordset")
  rs.Open "select * from [Class_lx] order by oId desc,cid desc",conn,1,1%>
<span class="clat">户型:</span><span class="TAB_CLICKa"><a href="?<%call seaurl(bc,sc,fj,0,lb,order,1)%>" <%call ison(lx,0)%>>不限</a>

<%do while not rs.eof%>
     <a href="?<%call seaurl(bc,sc,fj,rs("cid"),lb,order,1)%>" rel="nofollow" <%call ison(lx,rs("cid"))%>><%=rs("cname")%></a>
       <%rs.movenext
    loop
	rs.close
	set rs=nothing%>

</span><div class="c line1px"></div>
<%set rs=Server.CreateObject("ADODB.Recordset")
  rs.Open "select * from [Class_lb] order by oId desc,cid desc",conn,1,1%>
<span class="clat ">物业:</span><span class="TAB_CLICKa "><a href="?<%call seaurl(bc,sc,fj,lx,0,order,1)%>" <%call ison(lb,0)%>>不限</a>

<%do while not rs.eof%>
     <a href="?<%call seaurl(bc,sc,fj,lx,rs("cid"),order,1)%>" rel="nofollow" <%call ison(lb,rs("cid"))%>><%=rs("cname")%></a>
       <%rs.movenext

    loop
	rs.close
	set rs=nothing%>
</div>
<div class="cla2 container container2 ">
<span class="TAB_CLICKa"><a class="<%if order=0 then%>on<%end if%> hand" href="?<%call seaurl(bc,sc,fj,lx,lb,0,1)%>" rel="nofollow">热门推荐</a><a class="hand <%if order=1 then%>on<%end if%>" href="?<%call seaurl(bc,sc,fj,lx,lb,1,1)%>" rel="nofollow">总价</a><a class="hand <%if order=2 then%>on<%end if%>" href="?<%call seaurl(bc,sc,fj,lx,lb,2,1)%>" rel="nofollow">时间</a></span>
</div>
<div class="c"></div>
<%set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from [Table_Product] where Passed=true"
if bc<>0 and bc<>"" then sql=sql+" and bigclassid="&bc
if sc<>0 and sc<>"" then sql=sql+" and smallclassid="&sc
if fj<>0 and fj<>"" then sql=sql+" and cfjid='"&fj&"'"
if lx<>0 and lx<>"" then sql=sql+" and clxid like '%"&lx&"%'"
if lb<>0 and lb<>"" then sql=sql+" and clbid='"&lb&"'"
if key<>"" then sql=sql+" and (title like '%"&key&"%' or content like '%"&key&"%' or content1 like '%"&key&"%')"
 sql=sql+" and ckfsid='2'"
sql=sql+" order by OrderId desc,"
select case order
	case 0
	sql=sql+"gid desc,Elite,"
	case 1
	sql=sql+"jgzj desc,"
	case 2
	sql=sql+"jgjj desc,"
	case 3
	sql=sql+"updatetime desc,"
	case 4
	sql=sql+"jgzj,"
	case 5
	sql=sql+"jgjj,"
	case 6
	sql=sql+"updatetime,"
end select
sql=sql+"articleid desc"
rs.Open sql,conn,1,1%>
<!--part1-->
<div class="part1">
<div class="container container2">
    <div class="row mb20">

    <%
    if rs.bof and rs.eof then
    response.write("暂无内容")
  '  response.write sql
    else
    rs.PageSize=15 '设置页码
    pagecount=rs.PageCount '获取总页码
    page=int(page) '接收页码
    if page<=0 then page=1 '判断
    if request("page")="" then page=1
    rs.AbsolutePage=page '设置本页页码
    i=0
    do while not rs.eof and i<rs.PageSize
    %>
<!--start-->
        <div class="col-md-4 mt20">
        <a class="bor  tra" href="rentdetail.asp?id=<%=rs("articleid")%>">
                <div class="pic " style="background-image: url('<%=rs("defaultpicurl")%>');background-size:cover;">
                      <div class="pmask tra"><span>了解详情</span></div>
              </div>
                <div class="info">
                   <p class="tit"><%=rs("title")%></p>
                   <p class="price"><%if rs("jgzj")=0 then%>价格待定<%else%><%=rs("jgzj")%>万起<%end if%>/月</p>
                      <p class="txt">类型：<%call showName("class_lb",rs("clbid"),"cid","cname")%><br>区域：<%=rs("bigclassname")%>，<%=rs("smallclassname")%><br>
                   地址：<%x1=split(rs("Product_Id"),"|")%><%=x1(0)%></p>
                   
                </div>
                
              </a>
            </div>
<!--end-->
    
<%
rs.movenext
i=i+1
loop
end if
	rs.close
	set rs=nothing%>

                    </div>

 </ul>
</div>
</div>
<!--#include file="inc/footer.asp"-->