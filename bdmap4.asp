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
var point=new BMap.Point(110.803154,19.559595);
mp.centerAndZoom(point, 15);
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
this.getElementsByTagName("span")[0].innerHTML = that._overText;
arrow.style.background = "url(blue2.gif) no-repeat -8px -100px";
leftBar.style.background = "url(blue2.gif) no-repeat -12px -2px";
rightBar.style.background = "url(blue2.gif) no-repeat -22px -2px";
}
div.onmouseout = function(){
this.style.background = "url(blue.gif) repeat-x 0 -33px";
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
//var sContent ="<div class='q_MapInfo'><h4>楼盘标题</h4><div class='q_infos'><div class='q_infoLeft'><p>起价：11000元/平米</p><p>地址：文昌市清澜高隆湾旅游大道经纬花园</p><a href='#' target='_blank'>查看楼盘详细&gt;&gt;</a></div><div class='q_infoRight'><img src='images/a.jpg' /></div></div><div class='q_infoBot'><p>楼盘销售电话（售楼处工作时间：24小时）</p><em>0898-6666666</em></div></div>"
// var infoWindow = new BMap.InfoWindow(sContent); // 创建信息窗口对象

//下面开始尝试数组
var markerArr = [{title:"绿岛",content:'<div style="float:left;width:120px;">起价：6800 元/平米<br><br><span style="line-height:20px;display:inlin-block;">地址：文昌市清澜开发区文清大道北侧（椰子大观园对面）</span> <br><br><a href="/House-show?id=8">查看楼盘详细>></a></div><div style="float:right;width:80px;"><img src="/Uploads/House/201302/s_510cbb671095a.jpg" wdith="80px" height="75px"></div>',point:'110.804978|19.558234',isOpen:0,icon:{w:21,h:21,l:0,t:0,x:6,lb:5},fid:1},{title:"白金海岸",content:'<div style="float:left;width:120px;">起价：9000 元/平米<br><br><span style="line-height:20px;display:inlin-block;">地址：海南省文昌市清澜开发区高隆湾旅游大道一环西</span> <br><br><a href="/House-show?id=48">查看楼盘详细>></a></div><div style="float:right;width:80px;"><img src="/Uploads/House/201304/s_515eef07be8ff.jpg" wdith="80px" height="75px"></div>',point:'110.820752|19.53851',isOpen:0,icon:{w:21,h:21,l:0,t:0,x:6,lb:5},fid:2}]

for (var i = 0; i < markerArr.length; i++) {
var json=markerArr[i]
var txt = markerArr[i].title;
var fid = markerArr[i].fid;
var pintx = markerArr[i].point.split('|')[0];
var pinty = markerArr[i].point.split('|')[1];
mp.addOverlay(new ComplexCustomOverlay(new BMap.Point(pintx, pinty), txt, txt));

(function(){
var index = i;
var _iw = createInfoWindow(i);
//alert(mycom)
//var _marker =mycom;
//_marker.addEventListener("click",function(){
// mp.openInfoWindow(_iw);
// });
})()
}



  /*gavin code*/
  mp.addEventListener('click',function(){
console.log("#_#"+ fid);


  });

</script>