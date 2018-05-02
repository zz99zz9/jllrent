<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%
dim tdkid
tdkid=1

%>
<!--#include file="inc/tdk.Asp"-->
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/1.css"/>
    <style>
    .my-navbar,.leavetop.my-navbar{border-bottom:0;}
.ed1{    height: 560px;
        background-image: url("xgwl/img/1/bg01.png");
        background-size: cover;
        background-position: center top;
        position: relative;}
        .ed2{    height: 560px;
                background-image: url("xgwl/img/1/bg02.png");
                background-size: cover;
                background-position: center top;
                position: relative;}
                .ed3{    height: 560px;
                        background-image: url("xgwl/img/1/bg03.png");
                        background-size: cover;
                        background-position: center top;
                        position: relative;}
</style>
<!--广告部份-->
<div class="ed">
<div id="carousel-example-generic" class="carousel slide" data-ride="carousel">

        <ol class="carousel-indicators">
            <li data-target="#carousel-example-generic" data-slide-to="0" class="active"></li>
            <li data-target="#carousel-example-generic" data-slide-to="1"></li>
            <li data-target="#carousel-example-generic" data-slide-to="2"></li>
        </ol>


        <div class="carousel-inner" role="listbox">


<%Set prs= Server.CreateObject("ADODB.Recordset")
                                psql="select top 3 * From [links] where wid=1 order by oid desc,ID desc"
                                prs.open psql,conn,1,1
                                i=0%>
                                <%do while not prs.eof%>
                                <div class="item <%if i=0 then%>active<%end if%> ed1" style=" background-image: url('<%=prs("file")%>');" >

                                            </div>
                         <%
                        	i=i+1
                        	     prs.movenext
                        	loop
                        prs.close
                        set prs=nothing
                        %>

        </div>


<!--        <a class="left carousel-control" href="#carousel-example-generic" role="button" data-slide="prev">
                <span class="glyphicon glyphicon-chevron-left" aria-hidden="true"></span>
                <span class="sr-only">Previous</span>
            </a>
            <a class="right carousel-control" href="#carousel-example-generic" role="button" data-slide="next">
                <span class="glyphicon glyphicon-chevron-right" aria-hidden="true"></span>
                <span class="sr-only">Next</span>
            </a>-->
    </div>
<!--    <div class="edtxt wow fadeInUp">
<img src="xgwl/img/edtit.svg">
    </div>-->
    <div class="searchbar">
        <div class="searchint" style="display:block;"><input class="sint" value="华侨城苏河湾" type="text" id="skey"><button class="sbtn sxf" ><i class="mag"></i></button></div>
        <div class="searchint"><input class="sint" value="御品大厦" type="text"  id="okey"><button class="sbtn s2sf" ><i class="mag"></i></button></div>
    </div>
</div>
<!--part1-->
<link rel="stylesheet" href="xgwl/css/lib/swiper-4.1.0.min.css">
  <!-- Link Swiper's CSS -->
<div class="part1">
    <div class="container">
        <div class="row">
            <style>
     .swiper-container {
        width: 100%;
        height: 100%;
      }
      .swiper-slide {
        text-align: center;
        font-size: 18px;
        background: #fff;
  
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
    .swiper-container{margin-top:40px;margin-bottom:40px;}
    .swiper-container .swiper-slide .bor{border:1px solid #f6f6f6;width:100%;display:block;}
    .swiper-container .swiper-slide .bor:hover{border:1px solid #d6d5d5;}
    .pic{width:100%;height:260px;border-top:5px solid #d60808;}
    .pic .pmask{height:100%;width:100%;opacity: 0;line-height:260px;}
    .pic .pmask span{color:#fff;border:1px solid #fff;border-radius:5px;padding:5px 12px;}
    .swiper-container .swiper-slide .bor:hover .pic .pmask{background:rgba(255, 0, 0, 0.589);opacity: 1;}

           </style>
              <!-- Swiper -->
  <div class="swiper-container">
        <div class="swiper-wrapper">

<%set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from [Table_Product] where Passed=true and elite=true"
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

    <%
    if rs.bof and rs.eof then
    response.write("暂无内容")
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

            <!--循环开始-->
          <div class="swiper-slide">
              <a class="bor tra" href="rentdetail.asp?id=<%=rs("articleid")%>">
              <div class="pic " style="background-image: url('<%=rs("defaultpicurl")%>');background-size:cover;">
                    <div class="pmask tra"><span>了解详情</span></div>
            </div>
              <div class="info">
                 <p class="tit"><%=rs("title")%></p>
                 <p class="price"><%if rs("jgzj")=0 then%>价格待定<%else%><%=rs("jgzj")%>万起<%end if%>/月</p>
                    <p class="txt">类型：<%call showName("class_lb",rs("clbid"),"cid","cname")%><br>区域：<%=rs("bigclassname")%>，<%=rs("smallclassname")%><br>
                   地址：<%x1=split(rs("Product_Id"),"|")%><%=x1(0)%></p>
                 
              </div>
              
            </a>
          </div>
<!--循环结束-->
<!--end-->
    
<%
rs.movenext
i=i+1
loop
end if
	rs.close
	set rs=nothing%>


        </div>
        <!-- Add Pagination <div class="swiper-pagination"></div>-->
        
   <!-- Add Arrows -->
   <div class="swiper-button-next"></div>
   <div class="swiper-button-prev"></div>
      </div>
        <!-- Swiper JS -->
  <script src="xgwl/js/lib/swiper-4.1.0.min.js"></script>

  <!-- Initialize Swiper -->
  <script>
       var swiper = new Swiper('.swiper-container', {
      slidesPerView: 3,
      spaceBetween: 50,
      // init: false,
      pagination: {
        el: '.swiper-pagination',
        clickable: true,
      },
      breakpoints: {
        1024: {
          slidesPerView: 3,
          spaceBetween: 40,
        },
        768: {
          slidesPerView: 2,
          spaceBetween: 30,
        },
        640: {
          slidesPerView: 1,
          spaceBetween: 20,
        },
        320: {
          slidesPerView: 1,
          spaceBetween: 10,
        }
      },navigation: {
        nextEl: '.swiper-button-next',
        prevEl: '.swiper-button-prev',
      }
    });
    
  </script>
  <a class="morebtn" href="rent.asp">查看更多房源</a>
        </div>
    </div>
</div>
<!--part2-->


        <div class="part2">
        <p class="tit">在JLL上发布您的房源信息</p>
        <p class="txt">通过以下链接发布</p>
        
       
        <a href="for-landlord.asp" class="morebtn tffy">投放房源</a>
        </div>


<!--part3-->
<div class="part3">
    <div class="container">
        <div class="row">
            <p class="ptit">我们的服务流程</p>
            <div class="hidden-sm"><img src="xgwl/img/rent/Group.png" class="group"></div>

        </div>
    </div>
</div>
<!--part4-->
<div class="part4">
    <div class="container">
        <div class="row">
                <p class="ptit">荣获奖项</p>
                <div class="col-md-2 ppic col-xs-6"><img src="xgwl/img/rent/Asia-Pacific-Property-Awards-2016_0.png"></div>
                <div class="col-md-2 ppic col-xs-6"><img src="xgwl/img/rent/CR-100-Best-logo-2015_0.png"></div>
                <div class="col-md-2 ppic col-xs-6"><img src="xgwl/img/rent/WME_2017_color_0.png"></div>
                <div class="col-md-2 ppic col-xs-6"><img src="xgwl/img/rent/Fortune-World's-Most-Admired-Companies-2015.png"></div>
                <div class="col-md-2 ppic col-xs-6"><img src="xgwl/img/rent/euromoney_real_estate_2015_0.png"></div>
                <div class="col-md-2 ppic col-xs-6"><img src="xgwl/img/rent/500Logo_2016_0.png"></div>
        </div>
    </div>
</div>
<!--#include file="inc/footer.asp"-->
<script>
$(document).ready(function(){
    $(".navbar-brand img").attr("src","xgwl/img/logo.svg");
})
//搜索时input鼠标按下抬起效果

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