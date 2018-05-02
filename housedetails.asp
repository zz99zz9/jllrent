<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%id=request.QueryString("id")%>
<%set Rs=Server.CreateObject("ADODB.Recordset")
if id="" then
  sql="select * from [Table_Product]"
  else
  sql="select * from [Table_Product] where articleid="&id
  end if
'  sql=sql+"and SmallClassId="&mapid
  sql=sql+" order by Orderid"
 Rs.Open sql,conn,1,1%>
 <%dim t,d,k
 t=rs("wz")
 d=rs("Description")
 k=rs("Keywords")
 %>
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
    <link rel="stylesheet" href="xgwl/css/fyed.css"/>
<!--广告部份-->
<div class="navwz container">
<a href="http://www.jllresidential.cn">JLL</a> &gt; <a href="newhouse.asp"><%=rs("ckfsname")%></a> &gt; <a href="newhouse.asp?bc=<%=rs("bigclassid")%>"><%=rs("bigclassname")%></a> &gt; <a href="#"><%=rs("title")%></a>
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
    <div id="sildes" class='hidden'>
      <div class="control">
        <ul class="change">
        </ul>
      </div>
      <div class="thumbWrap">
        <div class="thumbCont">
          <ul>
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
    <!--新的轮播样式-->
    <style>
    .pc-slide {
	width: 700px;
	margin: 0 auto;
}

.view .swiper-container {
	width: 700px;
	height: 470px;
}
.view .swiper-container img{
	width: 100%;
  height:100%;

}
.view .arrow-left {
	background: url(xgwl/img/index_tab_l.png) no-repeat left top;
	position: absolute;
	left: 10px;
	top: 50%;
	margin-top: -25px;
	width: 28px;
	height: 51px;
	z-index: 10;
}

.view .arrow-right {
	background: url(xgwl/img/index_tab_r.png) no-repeat left bottom;
	position: absolute;
	right: 10px;
	top: 50%;
	margin-top: -25px;
	width: 28px;
	height: 51px;
	z-index: 10;
}

.preview {
	width: 100%;
	margin-top: 10px;
	position: relative;
}

.preview .swiper-container {
	width: 680px;
	height: 82px;
	margin-left: 35px;
}

.preview .swiper-slide {
	width: 87px;
	height: 82px;
}

.preview .slide6 {
	width: 82px;
}

.preview .arrow-left {
	background: url(xgwl/img/feel3.png) no-repeat left top;
	position: absolute;
	left: 10px;
	top: 50%;
	margin-top: -9px;
	width: 9px;
	height: 18px;
	z-index: 10;
}

.preview .arrow-right {
	background: url(xgwl/img/feel4.png) no-repeat left bottom;
	position: absolute;
	right: 10px;
	top: 50%;
	margin-top: -9px;
	width: 9px;
	height: 18px;
	z-index: 10;
}

.preview img {
	padding: 1px;
}

.preview .active-nav img {
	padding: 0;
	border: 1px solid #F00;
}
.smpic img{width:80px;height:80px;}
@media (max-width: 768px) {
    .pc-slide {
	width: 100%;
	margin: 0 auto;
}

.view .swiper-container {
	width: 100%;
	height: 220px;
}
.preview .swiper-container{
  width:75%;
}
}
</style>
 <div class="pc-slide">
	<div class="view">
		<div class="swiper-container">
			<a class="arrow-left" href="#"></a>
			<a class="arrow-right" href="#"></a>
			<div class="swiper-wrapper">
         <%Set prs= Server.CreateObject("ADODB.Recordset")
psql="select * From [Table_pic] where pro_id="&rs("ArticleID")&" order by oid desc,ID desc"
prs.open psql,conn,1,1%>
<%do while not prs.eof%>

        <div class="swiper-slide">
					<img src="<%=prs("pic_url")%>" alt="">
				</div>
 <%
	i=i+1

	     prs.movenext
	loop
prs.close
set prs=nothing
%>
			</div>
		</div>
	</div>
	<div class="preview">
		<a class="arrow-left" href="#"></a>
		<a class="arrow-right" href="#"></a>
		<div class="swiper-container">
			<div class="swiper-wrapper">
        <%Set prs= Server.CreateObject("ADODB.Recordset")
psql="select * From [Table_pic] where pro_id="&rs("ArticleID")&" order by oid desc,ID desc"
prs.open psql,conn,1,1
dim i
i=0%>
<%do while not prs.eof%>

        <div class="swiper-slide smpic <%if i=0 then%>active-nav<%end if%>">
					<img src="<%=prs("pic_url")%>" alt="">
				</div>
 <%
	i=i+1

	     prs.movenext
	loop
prs.close
set prs=nothing
%>

			</div>
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
    <div id="map" style="width:620px;height:420px;margin-bottom:50px;"></div>
</div>
<!---->
<!--原二维码和两个button位置-->


          <!---->
</div>
     <div class="pr wr">
     <h1><%=rs("title")%></h1>
     <span class="jiage">
		<%if rs("jgzj")=0 then%>
		价格待定
		<%else%>
		约¥<%=rs("jgzj")%><%if rs("jgjj")<>"" and rs("jgjj")<>"0" then%>-<%=rs("jgjj")%><%end if%>万起
	<%end if%></span>
     <span class="ntxt"><span class="ntxtl">区域<br>类别<br>户型<br><%if rs("mj1")<>"" then%>面积<%end if%></span><span class="ntxtr"><%=rs("bigclassname")%>，<%=rs("smallclassname")%><br><%=rs("clbname")%><br><%
                                                                                                                                                        mystr=split(rs("clxid"),",")
                                                                                                                                                          for i=0 to ubound(mystr)
                                                                                                                                                    if i<4 and mystr(i)<>"" then
                                                                                                                                                    %>
                                                                                                                                                    <em class="trait"><%call Showtd(mystr(i),"lx")%></em>
                                                                                                                                                    <% end if
                                                                                                                                                     next %><br><%if rs("mj1")<>"" then%><%=rs("mj1")%> <%if rs("mj2")<>"" and rs("mj2")<>0 then%>- <%=rs("mj2")%><%end if%> 平方米<%end if%></span>
      <div class="c"></div>
        <img src="xgwl/img/temp/bt1.png" class="bt1  b_h" onclick="fdjsq()"> <a href="news.asp?cid=11"><img src="xgwl/img/temp/bt2.png" class="bt2 b_h"></a>
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
                <a style="display:block;float:left;padding:5px 10px;background:#C00;color:#fff;font-size:12px;margin-left:240px;" href="newhouse.asp" target="_blank">更多</a><br /><br />
                <!--更多房源结 -->


	<!---->
	<div class="ltxtcc">
<div class="ltxt"><div class='ltxtinfo' ><img src="xgwl/img/qr.gif" style="margin-bottom:10px;"/><br />扫一扫 关注JLL</div></div>
    <div class="stitline" > <span class="stit">对此房源感兴趣？<a class="red" href="mailto:Sherry.liu@ap.jll.com">联系我们！</a></span>

      <div class="leftbtn b_h hidden-xs"><i></i><span>400-818-1133</span></div>
      <div class="rightbtn b_h hidden-xs" onClick="openfrom('20165new.pdf');"><i></i><span>留言咨询</span></div>
	  </div>
	</div>
	<!---->
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
<!--新轮播js-->

<script>
    var viewSwiper = new Swiper('.view .swiper-container', {
	onSlideChangeStart: function() {
		updateNavPosition()
	}
})

$('.view .arrow-left,.preview .arrow-left').on('click', function(e) {
	e.preventDefault()
	if (viewSwiper.activeIndex == 0) {
		viewSwiper.swipeTo(viewSwiper.slides.length - 1, 1000);
		return
	}
	viewSwiper.swipePrev()
})
$('.view .arrow-right,.preview .arrow-right').on('click', function(e) {
	e.preventDefault()
	if (viewSwiper.activeIndex == viewSwiper.slides.length - 1) {
		viewSwiper.swipeTo(0, 1000);
		return
	}
	viewSwiper.swipeNext()
})

var previewSwiper = new Swiper('.preview .swiper-container', {
	visibilityFullFit: true,
	slidesPerView: 'auto',
	onlyExternal: true,
	onSlideClick: function() {
		viewSwiper.swipeTo(previewSwiper.clickedSlideIndex)
	}
})

function updateNavPosition() {
		$('.preview .active-nav').removeClass('active-nav')
		var activeNav = $('.preview .swiper-slide').eq(viewSwiper.activeIndex).addClass('active-nav')
		if (!activeNav.hasClass('swiper-slide-visible')) {
			if (activeNav.index() > previewSwiper.activeIndex) {
				var thumbsPerNav = Math.floor(previewSwiper.width / activeNav.width()) - 1
				previewSwiper.swipeTo(activeNav.index() - thumbsPerNav)
			} else {
				previewSwiper.swipeTo(activeNav.index())
			}
		}
	}
</script>

