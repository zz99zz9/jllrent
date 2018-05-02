<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->

<!--#include file="inc/header.asp"-->
<!--row2:搜索条件-->

<div class="w">

  <div class="position"><a href="index.asp">JLL</a> &gt; <a href="privacy.asp">隐私条款</a>  </div>
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
.b{font-weight:bold;}
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
<div class="wl2" style="padding-left:10px;"><h1>隐私条款</h1></div>
<div class="wl2 rbor"><div class="w727"><div class="news-content">

<br>
仲量联行有限公司(纽约证交所交易代码：JLL)是专注于房地产领域的专业服务和投资管理公司，致力于为客户持有、租用和投资房地产的决策实现增值。仲量联行有限公司是《财富》500强上榜企业，业务遍及全球80个国家，拥有逾280个分公司，员工总数超过70,000人。2015年度业务营收达52亿美元，总收入60亿美元。2015年仲量联行代表客户管理和提供外包服务的物业总面积逾40亿平方英尺（约3.72亿平方米），并协助客户完成了价值1,380亿美元的物业出售、并购和融资交易。2016年前三季度，仲量联行旗下的投资管理业务分支“领盛投资管理（LaSalle Investment Management）”资产总值达597亿美元。
<br><br>
仲量联行有限公司在亚太地区开展业务超过50年。公司目前在亚太地区的16个国家拥有94个分公司，员工总数超过36,000人。在 “2016年国际物业奖”亚太区评选中，仲量联行囊括十五项大奖。此外，根据全球房地产交易量的独立机构Real Capital Analytics的数据显示，仲量联行连续五年蝉联亚太区房地产投资顾问公司榜首。
<br><br>
在大中华区，仲量联行有限公司目前拥有超过2,200名专业人员及14,000名驻场员工，所提供的专业房地产服务遍及全国80多个城市。在“2016年国际物业奖”评选中，仲量联行再度荣膺“中国最佳房地产咨询公司”，连续六年获此殊荣。仲量联行有限公司大中华区总部位于香港,地址：太古广场三期6楼，香港皇后大道东1号。联系电话：+852 58081883
<br><br>
<b class="b">我们对隐私的承诺</b><br>
我们认为，认真负责地使用在我们网站上收集的个人信息，对我们的业务目标和信誉是至关重要的。作为隐私承诺的一部分，我们采纳了这份隐私声明。从我们外部网站上收集到有关您的信息的程度和种类取决于您通过我们的邮箱和/或通过在我们网站上注册时提供的信息。
<br><br>
<b class="b">信息的使用</b><br>
我们只利用您的个人识别信息来回应您的要求，并向您提供关于仲量联行服务的信息。此外，我们还可以使用它来向您提供有关仲量联行服务的营销材料，除非您按照下文所述决定退出接收此类信息。同时，我们还将个人信息用于向客户提供服务、维持客户关系（例如，通知我们的客户和未来客户他们可能会感兴趣的房地产和服务）以及其它相关的业务活动之目的。
<br><br>
<b class="b">收集信息</b><br>
我们的某些网站会邀请您与我们联系或者提供有关您的个人信息。我们利用您提供的信息来个性化您的在线体验，并将最适合您需求的内容发送给您。
<br><br>
此外，在我们的某些网站上使用了网络跟踪器“cookies”（即，当您第一次访问本网站时记录在计算机中的一些小文本文件）。现在多数浏览器会在出现cookie时进行识别，并允许您表示拒绝或接受。如果您不肯定您的浏览器是否有这一功能，您应该同软件制造商、您公司的技术服务部门或者您的互联网服务供应商加以核实。
<br><br>
Cookies主要用来提升您的在线体验，而非用来跟踪相关访问者的上网习惯，除非我们获得您的许可。如果您访问我们的网站用于阅读或下载信息，例如新闻或文章，那么我们所收集的多数信息只是用于统计（例如，您接入互联网的域、访问我们网站的日期和时间、以及您直接连接到我们网站的互联网地址），而不是用于识别个人。我们利用关于访问者人数及其使用网站的信息加以汇总来使我们的网站对您更有用和更具吸引力。
<br><br>
<b class="b">共享信息</b><br>
您提供给我们的个人信息将被保留，仅用于实现信息收集或者法律要求的目的。我们不与第三方分享个人信息，除非在执行我们的业务或者按照您的请求或者依据法律或其它法律程序的要求，并且我们绝对不会出售您的个人信息。此外，在需要执行公司业务或者满足您提交信息的目的时，仲量联行收集的个人信息可以转给我们的分公司或附属公司。通过电子方式向仲量联行提交资料，这就明确表示您同意将通过我们网站所收集的信息进行跨境传播。
<br><br>
<b class="b">保护信息</b><br>
此外，我们拥有适当的技术和运营安全政策及程序，以防止您的个人识别信息丢失、滥用、更改或者肆意破坏。我们负责访问数据的工作人员已经接受了相关培训，以维护该信息的保密性。
<br><br>
仲量联行承认其具有保护您个人信息在线隐私的责任。如果您对我们在管理个人信息上存在疑问或意见，或者您想退出并不再从我们这里接收任何营销材料，请联络我们privacy@joneslanglasalle.com。您还可以使用这一电子邮件，就您关注的遵循本隐私政策方面的问题与我们交流。
<br><br>
<b class="b">同其它网站的链接</b><br>
本政策的使用范围包括仲量联行的公司网站。我们的部分分公司可能有其自己的、完全不同的政策。我们建议您在链接到这些网站时查看这些隐私政策。
<br><br>
<b class="b">地点与接受条款</b><br>
仲量联行在其位于美国伊利诺伊州芝加哥的办事处对本网站进行维护和运行。所有与本网站相关的事宜均受美国伊利诺伊州法律的管辖和解释。
<br><br>
通过访问本网站，表明您接受本隐私政策。
<br><br>
如有疑问请联系仲量联行:
<br><br>
地址：上海市静安区石门一路288号兴业太古汇1号楼22楼
<br><br>
联系电话:021-6133 5576

</div>
<div class="news-qr "><img src="xgwl/img/qr.gif" />扫一扫 关注JLL资讯</div>

<div class="clear"></div>
</div><div class="clear hidden-xs"></div>
</div>
<div class="wr2 ">

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
    <span class="wrtit">热点项目</span>
	<%set ors=Server.CreateObject("ADODB.Recordset")
osql="select top 5 * from [Table_Product] where Passed=true order by Elite,OrderId desc,articleid desc"

ors.Open osql,conn,1,1
i=0
do while not ors.eof%>
<a href="housedetails.asp?id=<%=ors("articleid")%>" target="_blank" style="color:#C00;line-height:35px;"><%=ors("title")%></a><br />
<a href="housedetails.asp?id=<%=ors("articleid")%>" target="_blank"><img src="<%=ors("defaultpicurl")%>" class="wgw b_h" /></a>
    <div class="wgwtxt fz12"><span class="b fz14"><%=ors("bigclassname")%>&nbsp;<%=ors("smallclassname")%></span>
      <div class="c"></div>
      约人民币 <span class="red"><%=ors("jgzj")%></span> 万<br />
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
%>
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