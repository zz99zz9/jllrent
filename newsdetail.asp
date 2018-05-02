<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%pro_id=request.QueryString("id")
call SafeRequest(pro_id)
set rsnew=Server.CreateObject("ADODB.Recordset")
sql="select * from [prod] where prod_id="&pro_id
rsnew.Open sql,conn,1,1
t=rsnew("prod_name")
if rsnew("keywords")<>"" then k=rsnew("keywords")
if rsnew("description")<>"" then d=rsnew("description")%>
<!--#include file="inc/header.asp"-->
<!--row2:搜索条件-->

<div class="w">

  <div class="position"><a href="http://www.jllresidential.cn">JLL</a> &gt; <a href="news.asp">热点资讯  &gt; <a href="#"><%=rsnew("classname")%>
  </a> </div>
</div>

<style>
.ltxt .smh{font-size:15px;font-weight:bold;color:#00F;}
.ltxt .smh:hover{color:#CD2927;}
.wl2 .gj li{padding:5px 12px;border-radius:5px;background-color:#f5f5f5;color:#333;float:left;margin-right:30px;margin-bottom:15px;font-size:14px;}
.wl2 .gj li a{color:#000;}
.wl2 .gj li:hover{background-color:#f5f5f5;color:#000;}
.wl2 .gj li:hover a{color:#000;}
.wl2 .gj li.on{background-color:#c00;color:#fff;}
.wl2 .gj li.on a{color:#fff;}
.c666{color:#666;}
.over280{overflow: hidden;text-overflow: ellipsis;white-space: nowrap;width:280px;}
.wr .bline,.wr2 .bline{border-bottom:1px #999999 dotted;height:20px;margin-bottom:20px;clear:both;width:285px;}
.news-more{margin-bottom:250px;}
.news-more span{width:50%;float:left;}
.news-more span a{color:#FE0000;}
.news-more span a:hover{color:#C00;}
.wl2 h1{font-size:22px;color:#000;margin-top:22px;}
.wl2 .news-date{color:#808080;font-size:10px;display:block;margin:16px 0;}
.wl2 .news-content{line-height:30px;}
.wl2 .news-content img{margin:0 0 0 0;max-width:100%;}/**/
.wl2 .news-qr{width:110px;font-size:12px;margin:50px 0 90px 265px;}
.wl2 .news-qr img{width:100%;margin-bottom:22px;}
.w727{width:727px;margin:0 auto 0 10px;clear:both;}
body{padding-top:100px;}
.w{width:1100px;margin:0 auto;clear:both;}
.position{color:#D32A27;line-height:70px;height:70px;font-size:12px;padding-left:10px;}
.position a{color:#000;}
.position a:hover{color:#666;}
.news-content a{color:#E30613;}
.wl2{width:780px;float:left;}
.wr2{width:300px;float:left;margin-left:20px;}
.wr .sbtn{width:130px;height:37px;background-color:#CD2927;padding-left:20px;line-height:37px;color:#fff;float:left;margin-right:15px;margin-bottom:15px;border:0px;background-repeat:no-repeat;background-position: 100px 8px;text-align:left;font-size:12px;}
.wr .ico_jsq{background-image:url(../img/ico_jsq.gif);}
.wr .ico_down{background-image:url(../img/ico_down.gif);}
.wr .ico_xm{background-image:url(../img/ico_xm.png);}
.wr .ico_js{background-image:url(../img/ico_js.png);}
.wr .wrtit,.wr2 .wrtit{display:block;clear:both;margin:30px 0px 25px 0px;position:relative;font-size:14px;font-weight:bold;}

.wr .wrint,.wr2 .wrint,.intpop .wrint{width:260px;height:45px;border-radius:10px;border:1px #ececec solid;background-color:#f2f2f2;line-height:50px;padding-left:10px;color:#4D4D4D;font-size:12px;}
.wr .errinfo,.wr2 .errinfo,.intpop .errinfo{display:block;height:20px;line-height:20px;font-size:12px;color:#C00;width:260px;}
.wr2 .wrtit{margin-top:0;}
.intpop{padding:30px 70px;display:none;}
.hlpop{display:none;width:610px;height:340px;margin-left:-305px;}
.intpop .wrtit{width:260px;display:block;text-align:center;margin-bottom:20px;}
.hltable{width:100%;height:100%;overflow:auto;}
.wr .wrbtn,.wr2 .wrbtn,.intpop .wrbtn{height:40px;width:130px;border:0px;background-color:#CD2927;color:#fff;text-align:center;line-height:40px;border-radius:6px;margin-bottom:20px;}
.intpop .wrbtn{width:260px;}
.wr .wgw,.wr2 .wgw{float:left;width:150px;height:100px;}
.wr .wgwtxt,.wr2 .wgwtxt{float:left;width:200px;line-height:200%;margin-left:15px;position:relative;height:100px;}
.rbor{border-right:1px solid #e6e6e6;}
.wr2 .wgwtxt{height:100px;width:130px;}
.wr .wgwtxt .more,.wr2 .wgwtxt .more{position:absolute;bottom:0;left:0;}
.wr .wgwtxt .more a,.wr2 .wgwtxt .more a{color:#000;}
.wr .wgwtxt .more a:hover,.wr2 .wgwtxt .more a:hover{color:#F00;}
@media (max-width: 768px) {
 .w,.wl2,.w727{width:100%;}
 .w727{margin:0 5px}
 .w{padding-right:10px;}
 .news-more span{width:100%;}
 .rbor{border-right:0;}
 .position{line-height:100%;height:20px;}
 .wl2 .news-qr{margin:50px auto;}
 .news-more{margin-bottom:120px;}
}
</style>
<!--row2:页面内容区-->
<div class="w">
<div class="wl2" style="padding-left:10px;"><h1><%=rsnew("prod_name")%></h1><span class="news-date"><%=formatdatetime(rsnew("prod_date"),2)%></span></div>
<div class="wl2 rbor"><div class="w727"><div class="news-content"><img src="<%=rsnew("prod_pic")%>" width="100%" title="<%=rsnew("picalt")%>" alt="<%=rsnew("picalt")%>"/><%prod_detail=Replace(rsnew("prod_detail"),"/jlladmin/kindeditor","/intl/jlladmin/kindeditor")%>
<%prod_detail=Replace(rsnew("prod_detail"),"/intl/intl/jlladmin/kindeditor","/intl/jlladmin/kindeditor")%>
<%=prod_detail%>
<div class="tips" style="font-size:14px;">
<%
if rsnew("tips")<>"" then
response.write"标签："
mystr=rsnew("tips")
mystr=replace(mystr,"，",",")
mystr=split(mystr,",")
for i=0 to ubound(mystr)
if i>0 then response.write" \ "
response.write "<a href='searchnews.asp?skey="&mystr(i)&"' style='color:#f00;font-size:14px;text-decoration:underline' target='_blank'>"&mystr(i)&"</a>"
next
end if  %>
</div>
</div>
<div class="news-qr "><img src="xgwl/img/qr.gif" />扫一扫 关注JLL资讯</div>
<div class="news-more"><%call sHowDxNewsPN(Rsnew("classid"),Rsnew("prod_id"))%></div>
<div class="clear"></div>
</div><div class="clear hidden-xs"></div>
</div>
<div class="wr2 ">
   <span class="wrtit">更多内容</span>
        <%set ors=Server.CreateObject("ADODB.Recordset")
osql="select top 5 * from [prod] where bigclassid="&rsnew("bigclassid")&" and classid=11 order by oid desc,prod_id desc"

ors.Open osql,conn,1,1
i=0
do while not ors.eof%>

    <div class="fz12 over280">
      <a href="newsdetail.asp?id=<%=ors("prod_id")%>" class=" c666" target="blank"><%=ors("prod_name")%></a></div>
      <div class="bline"></div>
    <%
ors.movenext
i=i+1
loop
	ors.close
	set ors=nothing%>
    <!---->
    <form>
      <span class="wrtit">免费咨询</span>
      <input type="text" class="wrint" value="您的姓名" id="wrint4"/><span class="errinfo"></span>
      <input type="text" class="wrint" value="手机号码" id="wrint5"/><span class="errinfo"></span>
      <input type="text" class="wrint" value="邮箱地址" id="wrint6"/><span class="errinfo"></span>
      <input type="button" class="wrbtn b_h" value="提交" id="wrbtn2"/>
    </form>
     <span class="wrtit">政策解读</span>
        <%set ors=Server.CreateObject("ADODB.Recordset")
osql="select top 5 * from [prod] where classid=12 or classid=11 or classid=15 order by oid desc,prod_id desc"

ors.Open osql,conn,1,1
i=0
do while not ors.eof%>

    <div class="fz12 over280">
      <a href="newsdetail.asp?id=<%=ors("prod_id")%>" class=" c666" target="blank"><%=ors("prod_name")%></a></div>
      <div class="bline"></div>
    <%
ors.movenext
i=i+1
loop
	ors.close
	set ors=nothing%>
    <span class="wrtit"><%=rsnew("bigclass")%>热点项目</span>
	<%set ors=Server.CreateObject("ADODB.Recordset")
osql="select top 5 * from [Table_Product] where Passed=true order by Elite,OrderId desc,articleid desc"

ors.Open osql,conn,1,1
i=0
do while not ors.eof%>
<a href="housedetails.asp?id=<%=ors("articleid")%>" target="_blank" style="color:#C00;line-height:35px;"><%=ors("title")%></a><br />
<a href="housedetails.asp?id=<%=ors("articleid")%>" target="_blank"><img src="<%=ors("defaultpicurl")%>" class="wgw b_h" /></a>
    <div class="wgwtxt fz12"><span class="b fz14"><%=ors("bigclassname")%>&nbsp;<%=ors("smallclassname")%></span>
      <div class="c"></div>
      <%if ors("jgzj")=0 then%>价格待定<%else%>
      约人民币 <span class="red"><%=ors("jgzj")%></span> 万<%end if%><br />
      <span class="more"><a href="newhouse.asp?id=<%=ors("articleid")%>">查看更多</a></span></div>
      <div class="bline"></div>
    <%
ors.movenext
i=i+1
loop
%>
    <a style="display:block;float:right;padding:5px 10px;background:#C00;color:#fff;font-size:12px;margin-right:15px;" href="6.asp" target="_blank">更多</a><br /><br />
    <%	ors.close
	set ors=nothing
		rsnew.close
	set rsnew=nothing%>
</div>
</div>
<!---->

<!--#include file="inc/footer.asp"-->
<script type="text/javascript" src="jll/js/jquery.cookies.js"></script>
<script type="text/javascript" src="jll/js/dfrom.js"></script>
<script type="text/javascript">

$("#nav4").attr("class","on");
//搜索时input鼠标按下抬起效果
input("#wrint1");
input("#wrint2");
input("#wrint3");
input("#wrint4");
input("#wrint5");
input("#wrint6");
input("#skey");
input("#okey");
function input(v){
    var n = $(v);
    var val=n.val();
    n.focus(function(){
        if(n.val()==val){
            $(this).val('');
            $(this).css("color","#000");
        }
    }).blur(function(){
        if (n.val()==""){
            $(this).val(val);
            $(this).css("color","#4d4d4d");
        }
    })
};
</script>