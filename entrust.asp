<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%
dim tdkid
tdkid=3
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
<div class="navwz container">
<a href="http://www.jllresidential.cn">JLL</a> &gt; <a href="entrust.asp">个人委托</a>
</div>
<!--广告部份-->
<div class="cla container"> <%set rs=Server.CreateObject("ADODB.Recordset")
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
<span class="clat">房价:</span><span class="TAB_CLICKa"><a href="?<%call seaurl(bc,sc,0,lx,lb,order,1)%>" <%call ison(fj,0)%>>不限</a>

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
<span class="clat">类别:</span><span class="TAB_CLICKa"><a href="?<%call seaurl(bc,sc,fj,lx,0,order,1)%>" <%call ison(lb,0)%>>不限</a>

<%do while not rs.eof%>
     <a href="?<%call seaurl(bc,sc,fj,lx,rs("cid"),order,1)%>" rel="nofollow" <%call ison(lb,rs("cid"))%>><%=rs("cname")%></a>
       <%rs.movenext

    loop
	rs.close
	set rs=nothing%>
</div>
<div class="cla2 container">
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
 sql=sql+" and ckfsid='1'"
sql=sql+" order by OrderId desc,"
select case order
	case 0
	sql=sql+"Elite,"
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
'response.write sql
rs.Open sql,conn,1,1%>
<!--part1-->
<div class="part1">
<div class="container">
    <ul class="fang">
    <%
    if rs.bof and rs.eof then
    response.write("暂无内容")
  '  response.write sql
    else
    rs.PageSize=10 '设置页码
    pagecount=rs.PageCount '获取总页码
    page=int(page) '接收页码
    if page<=0 then page=1 '判断
    if request("page")="" then page=1
    rs.AbsolutePage=page '设置本页页码
    i=0
    do while not rs.eof and i<rs.PageSize
    %>
    <li><a href="housedetails.asp?id=<%=rs("articleid")%>" class="fpic"><!--<i class="hot"></i>--><img src="<%=rs("defaultpicurl")%>" class="fangpic b_h"></a><div class="ninfo"><a  href="housedetails.asp?id=<%=rs("articleid")%>" class="ntit b_h"><%=rs("title")%></a><span class="ntxt"><span class="ntxtl">区域<br>户型<br>豪宅类型<br></span><span class="ntxtr"><%=rs("bigclassname")%>，<%=rs("smallclassname")%><br>
<%
    mystr=split(rs("clxid"),",")
      for i=0 to ubound(mystr)
if i<4 and mystr(i)<>"" then
%>
<em class="trait"><%call Showtd(mystr(i),"lx")%></em>
<% end if
 next %><br><%=rs("clbname")%><br></span>
    <div class="c"></div>
    <div class="lcl">
    <%
        mystr=split(rs("ctdid"),",")
          for i=0 to ubound(mystr)
    if i<4 and mystr(i)<>"" then
    %>
    <i class="cl<%=i%> "><%call Showtd(mystr(i),"td")%></i>

    <% end if
     next %>


    </div>
    </span>
    </div><div class="nphone"><span class="jiage">¥<%=rs("jgzj")%><b class="b1">万起</b><!--<b class="b2">(每平)</b>--></span><div class="c"></div><a href="housedetails.asp?id=<%=rs("articleid")%>" class="nbtn">查看更多</a></div></li>
<%
rs.movenext
i=i+1
loop
end if
	rs.close
	set rs=nothing%>

 </ul>
</div>
</div>

<!--#include file="inc/footer.asp"-->