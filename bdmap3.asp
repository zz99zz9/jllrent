<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css">
body, html,#allmap {width: 100%;height: 100%;overflow: hidden;margin:0;}
#l-map{height:100%;width:78%;float:left;border-right:2px solid #bcbcbc;}
#r-result{height:100%;width:20%;float:left;}
</style>
<script type="text/javascript" src="http://api.map.baidu.com/api?v=1.2&ak=CB2ede775afeb6e413abd40261396a69"></script>
<title>添加自定义覆盖物</title>
<%
mapid=request.QueryString("mapid")
if mapid<>"" then
set sRs=Server.CreateObject("ADODB.Recordset")
ssql="select * from [class_rq] where cid="&mapid
'response.write ssql
sRs.Open ssql,conn,1,1
center=sRs("cfile")
sRs.close
set sRs=nothing
if center="" then center="121.477288,31.224149"
'center="121.477288,31.224149"
else
center="121.477288,31.224149"
end if
%>

<style type="text/css">
*{margin:0;
padding:0;
}
.q_MapInfo{
overflow:hidden;
}

.q_MapInfo h4{
line-height:25px;
color:#2b2b2b;
border-bottom:1px solid #999;
font-size:12px;
margin-bottom:5px;
font-weight:normal;
}

.q_infos{
overflow:hidden;
}

.q_infos:after{
display:block;
visibility:hidden;
font-size:0;
height:0;
content:"";
clear:both;
}

*html .q_infos{zoom:1;}
*html+first-child .q_infos{zoom:1;}
.q_infoLeft{
float:left;
width:200px;
font-size:12px;
line-height:30px;
}
.q_infoLeft p{
line-height:30px;
}

.q_infoLeft a{
text-decoration:none;
color:#2b2b2b;
}
.q_infoLeft a:hover{
color:red;
}

.q_infoRight{
float:right;
width:100px;
height:75px;
}
.q_infoRight img{
display:block;
width:100px;
height:75px;
overflow:hidden;
}

.q_infoBot{
border-top:1px solid #999;
}

.q_infoBot p{
line-height:30px;
}

.q_infoBot em{
font-size:24px;
font-style:normal;
color:green;
}

</style>
</head>
<body>
<div id="allmap"></div>
</body>
</html>
<script type="text/javascript">
var mp = new BMap.Map("allmap");
mp.addControl(new BMap.OverviewMapControl()); //添加默认缩略地图控件
mp.addControl(new BMap.OverviewMapControl({isOpen:true, anchor: BMAP_ANCHOR_BOTTOM_RIGHT})); //右下角，打开
mp.addControl(new BMap.ScaleControl()); // 添加默认比例尺控件
mp.addControl(new BMap.ScaleControl({anchor: BMAP_ANCHOR_BOTTOM_LEFT})); // 左下
mp.addControl(new BMap.NavigationControl()); //添加默认缩放平移控件
mp.addControl(new BMap.NavigationControl({anchor: BMAP_ANCHOR_TOP_LEFT, type: BMAP_NAVIGATION_CONTROL_SMALL})); //右上角，仅包含平移和缩放按钮
var point=new BMap.Point(<%=center%>);

mp.centerAndZoom(point, 14);
mp.enableScrollWheelZoom();
// 复杂的自定义覆盖物
function ComplexCustomOverlay(point, text, mouseoverText){
this._point = point;
this._text = text;
this._overText = mouseoverText;
}
ComplexCustomOverlay.prototype = new BMap.Overlay();
ComplexCustomOverlay.prototype.initialize = function(map){
this._map = map;
var div = this._div = document.createElement("div");
div.style.position = "absolute";
div.style.zIndex = BMap.Overlay.getZIndex(this._point.lat);
div.style.background = "url(blue.gif) repeat-x 0 -33px";
div.style.color = "white";
div.style.height = "21px";
div.style.padding = "2px";
div.style.lineHeight = "18px";
div.style.whiteSpace = "nowrap";
div.style.MozUserSelect = "none";
div.style.fontSize = "12px"
var span = this._span = document.createElement("span");
div.appendChild(span);
span.appendChild(document.createTextNode(this._text));
var that = this;
var arrow = this._arrow = document.createElement("div");
arrow.style.background = "url(blue.gif) no-repeat -8px -100px";
arrow.style.position = "absolute";
arrow.style.width = "30px";
arrow.style.height = "12px";
arrow.style.top = "19px";
arrow.style.left = "10px";
arrow.style.overflow = "hidden";
div.appendChild(arrow);

var leftBar = this._leftBar = document.createElement("div");
leftBar.style.background = "url(blue.gif) no-repeat -12px -2px";
leftBar.style.position = "absolute";
leftBar.style.width = "11px";
leftBar.style.height = "24px";
leftBar.style.top = "0px";
leftBar.style.left = "-10px";
leftBar.style.overflow = "hidden";
div.appendChild(leftBar);
var rightBar = this._rightBar = document.createElement("div");
rightBar.style.background = "url(blue.gif) no-repeat -22px -2px";
rightBar.style.position = "absolute";
rightBar.style.width = "11px";
rightBar.style.height = "24px";
rightBar.style.top = "0px";
rightBar.style.right = "-10px";
rightBar.style.overflow = "hidden";
div.appendChild(rightBar);

div.onmouseover = function(){
this.style.background = "url(blue2.gif) repeat-x 0 -33px";
this.style.zIndex="1000";
this.getElementsByTagName("span")[0].innerHTML = that._overText;
arrow.style.background = "url(blue2.gif) no-repeat -8px -100px";
leftBar.style.background = "url(blue2.gif) no-repeat -12px -2px";
rightBar.style.background = "url(blue2.gif) no-repeat -22px -2px";
}
div.onmouseout = function(){
this.style.background = "url(blue.gif) repeat-x 0 -33px";
this.style.zIndex="100";
this.getElementsByTagName("span")[0].innerHTML = that._text;
arrow.style.background = "url(blue.gif) no-repeat -8px -100px";
leftBar.style.background = "url(blue.gif) no-repeat -12px -2px";
rightBar.style.background = "url(blue.gif) no-repeat -22px -2px";
}
mp.getPanes().labelPane.appendChild(div);

return div;
}

ComplexCustomOverlay.prototype.draw = function(){
var map = this._map;
var pixel = map.pointToOverlayPixel(this._point);
this._div.style.left = pixel.x - parseInt(this._arrow.style.left) + "px";
this._div.style.top = pixel.y - 30 + "px";
}

ComplexCustomOverlay.prototype.addEventListener = function(event,fun){//点击事件
this._div['on'+event] = fun;
}
//复杂物添加完后，我们该定义一下信息窗口了
function createInfoWindow(i){
var json = markerArr[i];
var iw = new BMap.InfoWindow("<b class='iw_poi_title' title='" + json.title + "'>" + json.title + "</b><p style=\"padding-top:8px;\"><hr></p><div class='iw_poi_content'>"+json.content+"</div>");
return iw;
}


//下面开始尝试数组
//var address=
<%set Rs=Server.CreateObject("ADODB.Recordset")
  sql="select * from [Table_Product] where ckfsid='2' "
'  sql=sql+"and SmallClassId="&mapid
  sql=sql+" order by Orderid"

 Rs.Open sql,conn,1,1
 i=0
  %>
var markerArr = [<%do while not rs.eof%><%if i>0 then%>,<%end if%>{title:"<%=rs("title")%>",content:'<div style="float:left;width:120px;">112</div>',point:'<%if rs("qt")="" then response.write("121.532911|31.249838") else response.write(rs("qt")) end if%>',fid:<%=rs("articleid")%>}<%rs.movenext
i=i+1
loop%>];
<%Rs.close
  set Rs=nothing%>


//var markerArr = [{title:"陆家嘴壹号院",content:'<div style="float:left;width:120px;">112</div>',point:'121.532911|31.249838',fid:1},{title:"九龙仓滨江壹十八",content:'<div style="float:left;width:120px;">222</div>',point:'121.522833|31.219519',fid:2},{title:"万科翡翠滨江",content:'<div style="float:left;width:120px;"></div>',point:'121.543602|31.250202',fid:3},{title:"仁恒公园世纪",content:'<div style="float:left;width:120px;"></div>',point:'121.541276|31.214027',fid:4},{title:"铂悦滨江",content:'<div style="float:left;width:120px;"></div>',point:'121.561477|31.248535',fid:5},{title:"浦东星河湾",content:'<div style="float:left;width:120px;"></div>',point:'121.542763|31.198567',fid:6},{title:"维诗凯亚",content:'<div style="float:left;width:120px;"></div>',point:'121.588351|31.239736',fid:7},{title:"凯旋滨江园",content:'<div style="float:left;width:120px;"></div>',point:'121.518493|31.227689',fid:15}];


for (var i = 0; i < markerArr.length; i++) {

var json=markerArr[i];
  var txt = markerArr[i].title;
  var fid = markerArr[i].fid;

  var pintx = markerArr[i].point.split('|')[0];
  var pinty = markerArr[i].point.split('|')[1];
  var marker = new Array();
  var point = new Array();
  point[i] = new window.BMap.Point(pintx, pinty);
  marker[i] = new window.ComplexCustomOverlay(point[i],txt,txt);
  mp.addOverlay(marker[i]);


(function(){
var index = i;
marker[i].addEventListener('click',function(e){
  parent.location.href="housedetails.asp?id="+markerArr[index].fid;
})



})()




}



</script>