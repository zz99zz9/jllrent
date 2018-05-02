<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%
dim tdkid
tdkid=1
%>
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
    <style>
    html{overflow: hidden;}
    iframe{overflow:hidden;width:100%;height:100%;}
</style>
<!--广告部份-->
<div class="serr container tra">
<div class="container ">
<div class="row">
<!--<img src="xgwl/img/temp/sear.gif">--><div class="col-md-2 "><select class="form-control"><option value="">区域</option><option value="">长宁区</option><option value="">黄浦区</option><option value="">静安区</option><option value="">浦东新区</option><option value="">宝山区</option><option value="">嘉定区</option><option value="">青浦区</option><option value="">卢湾区</option><option value="">松江区</option></select></div>
<div class="col-md-2"><select class="form-control"><option value="">房价</option><option value="">300-500万</option><option value="">500-700万</option><option value="">700-1000万</option><option value="">1000万以上</option></select></div>
<div class="col-md-2"><select class="form-control"><option value="">户型</option><option value="">1室1厅1卫</option><option value="">2室2厅2卫</option><option value="">3室2厅2卫</option><option value="">5室3厅3卫</option></select></div>
<div class="col-md-2"><select class="form-control"><option value="">类别</option><option value="">公寓</option><option value="">独栋别墅</option><option value="">联排别墅</option></select></div>
<div class="col-md-4 input-group"><input  class="form-control" type="text" placeholder="翠湖天地"><span class="input-group-btn"><button class="btn" type="button" onclick="location.href='4.asp'">搜索</button></span></div>
</div>
</div>
</div>
<!--广告部份-->

<!--part1-->
<div class="fanglist tra">
    <div class="jjia"><div class="tit">老西门<br><span>上海/黄浦</span></div><div class="rtxt">89061 <span>元/平</span></div></div>
    <div class="seain">为您找共 55 套在售房源<ul class="TAB_CLICK"><li class="on b_h">默认</li><li class=" b_h">最新</li><li class=" b_h">总价</li></ul></div>
    <div class="slist">
        <ul class="fang2"><li><img src="xgwl/img/f4.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        <li><img src="xgwl/img/f5.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        <li><img src="xgwl/img/f6.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        <li><img src="xgwl/img/f7.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        <li><img src="xgwl/img/f4.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        <li><img src="xgwl/img/f5.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        <li><img src="xgwl/img/f6.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        <li><img src="xgwl/img/f7.jpg"><span class="tit">宝格丽公寓</span><span class="txt">2室 2厅89.15平</span><span class="jiage2">1200万</span></li>
        </ul>


</div>

</div>
<div class="part7">
<a href="7.asp"><img src="xgwl/img/temp/mapp3.gif" class="mapp"></a>
<iframe scrolling='no'
		src='bdmap.asp'>
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