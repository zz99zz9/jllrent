<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
    <style>
    html{overflow: hidden;}
    body{background:#F3F3F4;}
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
<script>
function hello(h){
 $(".jjia .tit em").html(h);
 var dqdq=$(".jjia .tit em").html();
 if (dqdq=="卢湾"){
  location.href="8.asp?v=23";
 }else if (dqdq=="徐汇"){
   location.href="8.asp?v=24";
   }else if (dqdq=="华山路"){
         location.href="8.asp?v=26";
       }else if (dqdq=="南外滩"){
                location.href="8.asp?v=15";
               }else if (dqdq=="新外滩"){
                                  location.href="8.asp?v=27";
                               }else if (dqdq=="陆家嘴"){
                               //  $(".jjia .rtxt").html("96000 <span>元/平</span>");
                                 location.href="8.asp?v=29";
                               }else if (dqdq=="老西门"){
                                 location.href="8.asp?v=22";
                                  }else if (dqdq=="古北"){
                               location.href="8.asp?v=28";
                                }else if (dqdq=="静安"){
                              location.href="8.asp?v=25";
         }else if (dqdq=="虹桥"){
         location.href="8.asp?v=13";
         }else if (dqdq=="联洋"){
                   location.href="8.asp?v=37";
                   }else if (dqdq=="碧云"){
                    location.href="8.asp?v=30";
                    }else if (dqdq=="新江湾城"){
                                         location.href="8.asp?v=31";
                                         }else{
            location.href="8.asp?v=29";
          }



}
</script>
<a href="8.asp"><img src="xgwl/img/temp/mapp.gif" class="mapp"></a>
<ul class="mapp2"><a href="javascript:fd();"><i class="glyphicon glyphicon-plus"></i></a><a href="javascript:hy();"><i class="glyphicon glyphicon-refresh"></i></a><a href="javascript:sx();"><i class="glyphicon glyphicon-minus"></i></a></ul>

<!--<a href="javascript:fd();"><img src="xgwl/img/temp/mapp2.gif" class="mapp2"></a>-->
<script>
/**/function fd(){
  $(".mappp").css("width","1800px");
    $(".mappp").css("height","1800px");
  $(".mappp").css("margin-left","-233px");
  $(".mappp").css("margin-top","-350px");
  }
  function hy(){
    $(".mappp").css("width","1200px");
      $(".mappp").css("height","1200px");
    $(".mappp").css("margin-left","0px");
    $(".mappp").css("margin-top","0px");
    }
      function sx(){
        $(".mappp").css("width","800px");
          $(".mappp").css("height","800px");
        $(".mappp").css("margin-left","200px");
        $(".mappp").css("margin-top","350px");
        }
</script>
<div class="mappic" id="dialog"> <embed src="xgwl/swf/map.swf" width="1200" height="1200" class="mappp" wmode="transparent"></embed><!--<img src="xgwl/img/map.svg" class="mappic"><a ></a>--><div class="c"></div></div>
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
<script type="text/javascript">
    // 获取元素对象
    function g(id){return document.getElementById(id);}
    //自动居中
    function autoCenter(el){
        var bodyW=document.documentElement.clientWidth;
        var bodyH=document.documentElement.clientHeight;
        var elW=el.offsetWidth;
        var elH=el.offsetHeight;
        el.style.left=(bodyW-elW)/2+'px';
        el.style.top=(bodyH-elH)/2+'px';
    }
    //自动全屏
    function fillToBody(el){
        el.style.width=document.documentElement.clientWidth;
        el.style.width=document.documentElement.clientHeight;
    }

    var mouseOffsetX=0;//偏移
    var mouseOffsetY=0;
    var isDraging=false;//是否可拖拽
    //鼠标事件一 在标题栏上按下（要计算鼠标相对拖拽元素的左上角坐标，并且标记元素为可拖动）
    g('dialog').addEventListener('mousedown',function(e){
    console.log("*_*")
        var e=e || window.event;
        mouseOffsetX= e.pageX-g('dialog').offsetLeft;
        mouseOffsetY= e.pageY-g('dialog').offsetTop;
        isDraging=true;

    })
    //鼠标事件二 鼠标移动时（要检测 ，元素是否可标记为移动，如果是，则更新元素位置到当前鼠标位置）
    document.onmousemove=function(e){
        var e= e||window.event;
        var mouseX= e.pageX;
        var mouseY= e.pageY;
        var moveX=0;
        var moveY=0;
        if(isDraging==true){
            moveX = mouseX-mouseOffsetX;
            moveY= mouseY-mouseOffsetY;
            //范围限定
       /*     var pageWidth=document.documentElement.clientWidth;
            var pageHeight=document.documentElement.clientHeight;

            var dialogWidth=g('dialog').offsetWidth;
            var dialogHeight=g('dialog').offsetHeight;

            var maxX=pageWidth-dialogWidth;
            var maxY=pageHeight-dialogHeight;

            moveX=Math.min(maxX,Math.max(0,moveX));
            moveY=Math.min(maxY,Math.max(0,moveY));*/


            g('dialog').style.left=moveX+'px';
            g('dialog').style.top=moveY+'px';
        }
    }
    //鼠标事件三 鼠标松开时（标记元素为不可拖动）
    document.onmouseup=function(){
        isDraging=false;
    }
    function showDialog(){
        g('dialog').style.display='block';
        g('mask').style.display='block';
        autoCenter(g('dialog'));
        fillToBody(g('mask'));
    }
    function hideDialog(){
        g('dialog').style.display='none';
        g('mask').style.display='none';

    }
    window.onresize = function(){
        autoCenter(g('dialog'));
        fillToBody(g('mask'));
    }
</script>
<!---->
</body>
</html>