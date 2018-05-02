<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%
dim tdkid
tdkid=1
mapid=request.QueryString("v")
%>
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
    <style>
    html{overflow: hidden;}
    iframe{overflow:hidden;width:100%;height:100%;}
</style>
<!--广告部份-->
<div class="serr container tra ">
<div class="container ">
<div class="row">
<!--<img src="xgwl/img/temp/sear.gif">--><div class="col-md-2 "><select class="form-control"><option value="">区域</option>
  <%
set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from [Table_ProSmallclass] where bigclassid=34 order by orderid desc,smallclassid desc"
rs.open sql,conn,1,1

if not(rs.eof and rs.bof) then

do while not rs.eof
%>
            <option value="<%=rs("smallclassname")%>;<%=rs("smallclassid")%>"><%=rs("smallclassname")%></option>
            <%
rs.movenext
loop
end if
rs.close

%>
</select></div>
<div class="col-md-2"><select class="form-control"><option value="">房价</option>
  <%

sql="select * from [class_fj] order by oid desc"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then

do while not rs.eof
%>
            <option value="<%=rs("CName")%>;<%=rs("Cid")%>"><%=rs("CName")%></option>
            <%
rs.movenext
loop
end if
rs.close

%>
</select></div>
<div class="col-md-2"><select class="form-control"><option value="">户型</option>
 <%

sql="select * from [class_lx] order by oid desc"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then

do while not rs.eof
%>
            <option value="<%=rs("CName")%>;<%=rs("Cid")%>"><%=rs("CName")%></option>
            <%
rs.movenext
loop
end if
rs.close

%>
</select></div>
<div class="col-md-2"><select class="form-control"><option value="">类别</option>
 <%

sql="select * from [class_lb] order by oid desc"
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then

do while not rs.eof
%>
            <option value="<%=rs("CName")%>;<%=rs("Cid")%>"><%=rs("CName")%></option>
            <%
rs.movenext
loop
end if
rs.close

%>
</select></div>
<div class="col-md-4 input-group"><input  class="form-control" type="text" placeholder="翠湖天地"><span class="input-group-btn"><button class="btn" type="button" onclick="location.href='4.asp'">搜索</button></span></div>
</div>
</div>
</div>
<!--广告部份-->

<!--part1-->
<div class="fanglist tra">
    <div class="jjia"><div class="tit"><em>上海</em><br><span></span></div><!--<div class="rtxt">80000 <span>元/平</span></div>--></div>
    <div class="seain">为您找共 4 套在售房源<!--<ul class="TAB_CLICK"><li class="on b_h">默认</li><li class=" b_h">最新</li><li class=" b_h">总价</li></ul>--></div>
    <div class="slist">
        <ul class="fang2">
<%
dim i
dim sql,rs

sql="select * From [Table_Product] where ArticleID>0"
sql=sql & "and bigclassid=34 order by orderid desc,ArticleID desc"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
 <%do while not rs.eof%>
        <li class=" b_c tra" onclick="location.href='6.asp?id=<%=rs("articleid")%>'"><img src="<%=rs("defaultpicurl")%>"><span class="tit"><%=rs("title")%></span><span class="txt"><%=rs("bigclassname")%>，<%=rs("smallclassname")%></span><span class="jiage2"><%=rs("jgzj")%>万</span></li>

<%rs.movenext

    loop
	rs.close
	set rs=nothing%>
        </ul>


</div>

</div>
<div class="part7">
<a href="7.asp"><img src="xgwl/img/temp/mapp3.gif" class="mapp"></a>
<iframe scrolling='no'
		src='bdmap3.asp?mapid=<%=mapid%>'>
		</iframe>

</div>

<!------>
<script src="xgwl/js/lib/jquery.js"></script>
<script src="xgwl/js/lib/bootstrap.min.js"></script>
<script src="xgwl/js/lib/wow.min.js"></script>

<script>
    if (!(/msie [6|7|8|9]/i.test(navigator.userAgent))){
        new WOW().init();
    };
</script>

<script src="xgwl/js/base.js"></script>
<script>
$(document).ready(function(){
    $(".navbar-brand img").attr("src","xgwl/img/logo3.svg");
    $(".navbar").addClass("leavetop");
     $(".serr").addClass("seer2");
                 $(".fanglist").addClass("f2");

})
</script>
<!---->

<!---->
</body>
</html>