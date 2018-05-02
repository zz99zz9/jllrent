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
 t=rs("entitle")
 d=rs("Description")
 k=rs("Keywords")
 %>
<!--#include file="inc/enheader.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
        <style>
        .ninfo .ntxt{border-top:0px;}
    </style>
<!--广告部份-->
<style>
.txt.en{font-size:1em;}
</style>
<!--广告部份-->
<div class="part1 ppi">
        <div class="container container2">
                <div class="finfo">
                        <p class="tit en">
                            <%=rs("entitle")%>
                            <span class="price en"><%=rs("jgzj")*10%><%if rs("jgjj")<>"" then%>-<%=rs("jgjj")*10%><%end if%>k/monthly</span>
                        </p>
                        <p class="txt en">
                                <%x1=split(rs("Product_Id"),"|")%><%=x1(1)%> <span class="right"><%call showName("class_lb",rs("clbid"),"cid","cnameen")%></span>
                        </p>
                        <a href="tel:400-8181-133" class="fbtn">Contact us</a>
                    </div>
            <div class="row">
<!-- Link Swiper's CSS -->
<link rel="stylesheet" href="xgwl/css/lib/swiper-4.1.0.min.css">

<!-- Demo styles -->
<style>
  .swiper-container {
    width: 100%;
    height: 100%;
    margin-left: auto;
    margin-right: auto;
  }
  .swiper-slide {
    text-align: center;
    font-size: 18px;
    background: #fff;
border-top:5px solid #d60808;
    /* Center slide text vertically */
    display: -webkit-box;
    display: -ms-flexbox;
    display: -webkit-flex;
    display: flex;
    -webkit-box-pack: center;
    -ms-flex-pack: center;
    -webkit-justify-content: center;
    justify-content: center;
    -webkit-box-align: center;
    -ms-flex-align: center;
    -webkit-align-items: center;
    align-items: center;
  }
  .swiper-slide img{width:100%;border-top:5px solid #d60808;}

</style>

<!-- Swiper -->
<div class="swiper-container">
  <div class="swiper-wrapper">
   <%Set prs= Server.CreateObject("ADODB.Recordset")
psql="select * From [Table_pic] where pro_id="&rs("ArticleID")&" order by oid desc,ID desc"
prs.open psql,conn,1,1%>
<%do while not prs.eof%>

                <div class="swiper-slide"><img src="<%=prs("pic_url")%>"></div>
 <%
	i=i+1

	     prs.movenext
	loop
prs.close
set prs=nothing
%>



  </div>
  <!-- Add Pagination -->
  <div class="swiper-pagination"></div>
  <!-- Add Arrows -->
  <div class="swiper-button-next"></div>
  <div class="swiper-button-prev"></div>
</div>

<!-- Swiper JS -->
<script src="xgwl/js/lib/swiper-4.1.0.min.js"></script>

<!-- Initialize Swiper -->
<script>
  var swiper = new Swiper('.swiper-container', {
    slidesPerView: 1,
    spaceBetween: 30,
    loop: true,
    pagination: {
      el: '.swiper-pagination',
      clickable: true,
    },
    navigation: {
      nextEl: '.swiper-button-next',
      prevEl: '.swiper-button-prev',
    },
  });
</script>
</div>
</div>
</div>
<!--part1-->
<div class="part1">
        <div class="container container2 ">
            <div class="row mb20">
                <p class="ptit">Information</p>
                <p class="ptxt"><%=rs("content1")%>
                </p>
            </div>
        </div>
</div>
<!--part1-->
<div class="part1">
        <div class="container container2 ">
            <div class="row mb20">
                <p class="ptit">Map</p>
                <p class="ptxt"><!--<img src="xgwl/img/rent/bitmap2.png">-->
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
    <div id="map" style="margin-bottom:50px;"></div>
                </p>
            </div>
        </div>
</div>
<!--part1-->
<div class="part1">
<div class="container container2">

    <div class="row mb20">
            <p class="ptit">MORE</p>
<%set rs=Server.CreateObject("ADODB.Recordset")
sql="select top 6 * from [Table_Product] where Passed=true"
if bc<>0 and bc<>"" then sql=sql+" and bigclassid="&bc
if sc<>0 and sc<>"" then sql=sql+" and smallclassid="&sc
if fj<>0 and fj<>"" then sql=sql+" and cfjid='"&fj&"'"
if lx<>0 and lx<>"" then sql=sql+" and clxid like '%"&lx&"%'"
if lb<>0 and lb<>"" then sql=sql+" and clbid='"&lb&"'"
if key<>"" then sql=sql+" and (title like '%"&key&"%' or content like '%"&key&"%' or content1 like '%"&key&"%')"
 sql=sql+" and ckfsid='2'"
sql=sql+" order by OrderId desc,"
select case order
	case 0
	sql=sql+"gid desc,Elite,"
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
rs.Open sql,conn,1,1%>
<!--part1-->
<div class="part1">
<div class="container container2">
    <div class="row mb20">

    <%
    if rs.bof and rs.eof then
    response.write("nothing")
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
<!--start-->
        <div class="col-md-4 mt20">
        <a class="bor  tra" href="enrentdetail.asp?id=<%=rs("articleid")%>">
                <div class="pic " style="background-image: url('<%=rs("defaultpicurl")%>');background-size:cover;">
                      <div class="pmask tra"><span>Detail</span></div>
              </div>
                <div class="info">
                   <p class="tit"><%=rs("entitle")%></p>
                   <p class="price">Rant:<%if rs("jgzj")=0 then%>undetermined<%else%>¥<%=rs("jgzj")%>k+<%end if%></p>
                      <p class="txt">Type：<%call showName("class_lb",rs("clbid"),"cid","cnameen")%><br>Area：<%call showName("Table_ProBigClass",rs("bigclassid"),"bigclassid","bigclassnameen")%>，<%call showName("Table_ProSmallClass",rs("smallclassid"),"smallclassid","smallclassnameen")%><br>
                   Address：<%x1=split(rs("Product_Id"),"|")%><%=x1(1)%></p>
                   
                </div>
                
              </a>
            </div>
<!--end-->
    
<%
rs.movenext
i=i+1
loop
end if
	rs.close
	set rs=nothing%>


                    </div>

    
</div>
</div>
<!--#include file="inc/enfooter.asp"-->