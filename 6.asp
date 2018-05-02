<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%id=request.QueryString("id")
%>
<%set Rs=Server.CreateObject("ADODB.Recordset")
if id="" then
  sql="select * from [Table_Product]"
  else
  sql="select * from [Table_Product] where articleid="&id
  end if
'  sql=sql+"and SmallClassId="&mapid
  sql=sql+" order by Orderid"
 Rs.Open sql,conn,1,1%>
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
    <link rel="stylesheet" href="xgwl/css/fyed.css"/>
<!--广告部份-->
<div class="navwz container">
<a href="/guonei">JLL</a> &gt; <a href="4.asp"><%=rs("ckfsname")%></a> &gt; <a href="4.asp?bc=<%=rs("bigclassid")%>"><%=rs("bigclassname")%></a> &gt; <a href="#"><%=rs("title")%></a>
</div>


<!--part1-->
<div class="part1">
<div class="container" >
    <div class="pl wl">
    <!------广告轮播----->
<!--轮播开始-->
<style>
#sildes{margin-left:-15px !important;}
</style>
    <div id="sildes" >
      <div class="control">
        <ul class="change">
        </ul>
      </div>
      <div class="thumbWrap">
        <div class="thumbCont">
          <ul>
            <!-- img属性, url=url, text=描述, bigimg=大图, alt=标题   url="http://www.sucaijiayuan.com"-->
 <%Set prs= Server.CreateObject("ADODB.Recordset")
psql="select * From [Table_pic] where pro_id="&rs("ArticleID")&" order by oid desc,ID desc"
prs.open psql,conn,1,1%>
<%do while not prs.eof%>
            <li>
              <div><img src="<%=prs("pic_url")%>" bigImg="<%=prs("pic_url")%>" ></div>
            </li>
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
    </div>
    <!--轮播结束-->
<%if rs("gid")<>0 and rs("gid")<>"" then%>
 <%Set prs= Server.CreateObject("ADODB.Recordset")
psql="select * From [team] where id="&rs("gid")&" order by oid desc,ID desc"
prs.open psql,conn,1,1%>
<div class="people"><img class="pic" src="<%=prs("pic")%>" height="100px"><div class="text"><div class="tit"><%=prs("cname")%> <%=prs("ename")%> <span class='hidden-xs'>&nbsp;&nbsp;<%=prs("tel")%><br><%=prs("position")%></span></div><div class="txt"><%=prs("info")%></div></div><a href="tel:<%=prs("tel")%>"><img class="telbtn" src="xgwl/img/temp/telbtn.gif"></a></div>
<%prs.close
set prs=nothing
%>
<%end if%>
<ul class="nav2 "><li class="on hand  tra"  ><a href="#f1" class="topLink">项目简介</a></li><li class="hand"><a href="#f2" class="topLink">户型图</a></li><li class="hand"><a href="#f3" class="topLink">地图</a></li></ul>
<style>
h2{font-size:15px;font-weight:700;padding-top:20px;position:relative;}
h2 .pao{position:absolute;top:-100px;left:0px;}
@media (max-width: 768px) {
.wl .stitline{margin:30px auto;}
.pr{width:100%;}
.wr .wgwtxt, .wr2 .wgwtxt{width:150px;}
.wr .wgw, .wr2 .wgw{width:120px;height:80px;}
.pr .jiage{font-size:20px;}
}
</style>
<!--项目简介-->
<h2>项目简介<div class="pao" id="f1"></div></h2>
<div class="">
<%=rs("content1")%>
</div>
<!--周边设施-->
<h2>户型图<div class="pao" id="f2"></div></h2>
<div class="">
<%=rs("content")%>
</div>
<!--地图-->
<h2>地理位置<div class="pao" id="f3"></div></h2>
<div class="">
 <!-- 百度地图API-->
 <script src="http://api.map.baidu.com/api?v=1.2" type="text/javascript"></script>
    <script type="text/javascript">
    function initialize() {
        //创建地图实例
        var map = new BMap.Map('map');
        //创建一个坐标
        var point =new BMap.Point(<%=Replace(rs("qt"),"|",",")%>);
        //地图初始化，设置中心点坐标和地图级别
        map.centerAndZoom(point,15);
        map.addControl(new BMap.MapTypeControl());
        map.addControl(new BMap.NavigationControl());
        var marker = new BMap.Marker(point);  　　　　　 // 创建标注
        map.addOverlay(marker);                        // 将标注添加到地图中
    }
    window.onload = initialize;
    </script>
<!-- 百度地图地图容器-->
    <div id="map" style="width:620px;height:420px"></div>
</div>
<!---->
<div class="ltxt"><div class='ltxtinfo' ><img src="xgwl/img/qr.gif" style="margin-bottom:10px;"/><br />扫一扫 关注JLL</div></div>
    <div class="stitline" > <span class="stit">对此房源感兴趣？<a class="red" href="mailto:Sherry.liu@ap.jll.com">联系我们！</a></span>

      <div class="leftbtn b_h hidden-xs"><i></i><span>400-818-1133</span></div>
      <div class="rightbtn b_h hidden-xs" onClick="openfrom('20165new.pdf');"><i></i><span>留言咨询</span></div>
      </div>


          <!---->
</div>
     <div class="pr wr">
     <span class="tit"><%=rs("title")%></span>
     <span class="jiage">约¥<%=rs("jgzj")%><%if rs("jgjj")<>"" and rs("jgjj")<>"0" then%>-<%=rs("jgjj")%><%end if%>万起</span>
     <span class="ntxt"><span class="ntxtl">区域<br>类别<br>户型<br></span><span class="ntxtr"><%=rs("bigclassname")%>，<%=rs("smallclassname")%><br><%=rs("clbname")%><br><%
                                                                                                                                                        mystr=split(rs("clxid"),",")
                                                                                                                                                          for i=0 to ubound(mystr)
                                                                                                                                                    if i<4 and mystr(i)<>"" then
                                                                                                                                                    %>
                                                                                                                                                    <em class="trait"><%call Showtd(mystr(i),"lx")%></em>
                                                                                                                                                    <% end if
                                                                                                                                                     next %><br></span>
      <div class="c"></div>
        <img src="xgwl/img/temp/bt1.png" class="bt1  b_h" onclick="fdjsq()"> <a href="2.asp"><img src="xgwl/img/temp/bt2.png" class="bt2 b_h"></a>
            <div class="c"></div>
            <form>
              <span class="wrtit">免费咨询</span>
              <input type="text" class="wrint" value="您的姓名" id="wrint4"/><span class="errinfo"></span>
              <input type="text" class="wrint" value="手机号码" id="wrint5"/><span class="errinfo"></span>
              <input type="text" class="wrint" value="邮箱地址" id="wrint6"/><span class="errinfo"></span>
              <input type="button" class="wrbtn b_h" value="提交" id="wrbtn2"/>
            </form>
            <div class="c"></div>
            <!--更多房源-->
                <span class="wrtit">更多房源</span>
                    <%set ors=Server.CreateObject("ADODB.Recordset")
            osql="select top 5 * from [Table_Product] where Passed=true order by Elite,OrderId desc,articleid desc"'and Elite=true

            ors.Open osql,conn,1,1
            i=0
            do while not ors.eof%>
            <a href="?id=<%=ors("articleid")%>" target="_blank" style="color:#C00;line-height:35px;"><%=ors("entitle")%></a><br />
            <a href="?id=<%=ors("articleid")%>" target="_blank"><img src="<%=ors("defaultpicurl")%>" class="wgw b_h" /></a>
                <div class="wgwtxt fz12"><span class="b fz14"><%=ors("bigclassname")%>&nbsp;<%=ors("smallclassname")%></span>
                  <div class="c"></div>
                  约人民币 <span class="red"><%=ors("jgzj")%></span> 万+<br />
                  <span class="more"><a href="?id=<%=ors("articleid")%>" target="blank">查看更多</a></span></div>
                  <div class="bline"></div>
                <%
            ors.movenext
            i=i+1
            loop
            	ors.close
            	set ors=nothing%>
                <a style="display:block;float:left;padding:5px 10px;background:#C00;color:#fff;font-size:12px;margin-left:240px;" href="4.asp" target="_blank">更多</a><br /><br />
                <!--更多房源结 -->

</div>
     <div class="c"></div>


</div>
</div>
<!--房代计算器-->

<div class="fdjsq">
<span class="popclose b_c ">×</span>
<h1>房贷计算器</h1>
  <iframe src="calculator.html" width="100%" height="500" frameborder="0">
3        您的浏览器不支持iframe，请升级
4   </iframe>
</div>

<!--#include file="inc/footer.asp"-->
<script>

function fdjsq(){
$(".mask").fadeIn();
$(".fdjsq").fadeIn();
}

$(".popclose").click(function(){
$(".fdjsq").fadeOut();
$(".mask").fadeOut();
console.log("#_!");
});
</script>
<script type="text/javascript">
$(function(){
	$('#sildes').olvSlides(
		{
			thumb:true,
			thumbPage:true,
			thumbDirection:"Y",
			effect:'fade'

		}
	);
})
</script>

