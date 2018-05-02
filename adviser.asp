<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->

<%
dim tdkid
tdkid=6
%>
<!--#include file="./inc/tdk.asp"-->
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
    <style>.nphone {    background:url("xgwl/img/tel.png") no-repeat;}</style>
<!--广告部份-->
<div class="navwz container">
<a href="http://www.jllresidential.cn">JLL</a> &gt; <a href="adviser.asp">豪宅顾问</a>
</div>
<!--广告部份--
<div class="cla container">
<span class="clat">区域:</span><span class="TAB_CLICKa"><a class="on">不限</a><a >静安区</a><a >黄埔区</a><a >长宁区</a><a >松江区</a><a >嘉定区</a><a >闵行区</a><a >杨浦区</a></span>
</div>
<div class="cla2 container">
<span class="TAB_CLICKa"><a class="on hand">从来时间</a><a class="hand">职位排序</a></span>
</div>-->
<div class="c"></div>
<!--part1-->
<style>
.ntxtgw{width:350px;}
</style>
<div class="part1">
<div class="container">
    <ul class="news">
     <%Set prs= Server.CreateObject("ADODB.Recordset")
    psql="select * From [team] where cid=1 order by oid desc,ID desc"

    prs.open psql,conn,1,1%>
    <%do while not prs.eof%>

                 <li class="col-md-6"><div class="inbor"><img src="<%=prs("pic")%>"><div class="ninfo"><span class="ntit"><%=prs("cname")%> <em><%=prs("ename")%></em></span><span class="ntime"><%=prs("position")%></span><span class="ntxt ntxtgw"><br><a class="ntel" href="tel:<%=prs("tel")%>"><%=prs("tel")%></a><br><%=prs("summary")%><br><a href="mailto:<%=prs("mail")%>" class="mailaddress"><%=prs("mail")%></a></span></div><a href="javascript:openfrom('20165new.pdf');" class="nbtn hidden-xs">立即咨询</a></div></li>
     <%
    	i=i+1

    	     prs.movenext
    	loop
    prs.close
    set prs=nothing
    %>



 </ul>
</div>
</div>

<!--#include file="inc/footer.asp"-->