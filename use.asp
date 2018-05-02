<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->

<!--#include file="inc/header.asp"-->
<!--row2:搜索条件-->

<div class="w">

  <div class="position"><a href="index.asp">JLL</a> &gt; <a href="use.asp">使用条款</a>  </div>
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
<div class="wl2" style="padding-left:10px;"><h1>请在使用本网站之前认真阅读以下使用条件和免责声明</h1></div>
<div class="wl2 rbor"><div class="w727"><div class="news-content">

<br>
<b class="b">接受条款</b><br>
仲量联行有限公司（“仲量联行”）负责维护本网站（“网站”）供您个人使用。访问和使用本网站需要遵循以下使用条款：仲量联行保留随时更新使用条款的权利，而无需另行通知。最新版本的使用条款可以通过点击位于网站底部的“使用条款”超文本链接查看。在使用本网站时，您需要无条件接受这些使用条款。如果不同意这些使用条款，请不要使用本网站。
<br><br>
<b class="b">信息的准确和完整性</b><br>

仲量联行力争确保本网站包含的信息准确可靠，仲量联行就信息的准确、正确、可靠与否不做任何保证或陈述，并且对本网站内容中的遗漏或错误不承担任何义务或责任。
<br><br>
<b class="b">网站的修改</b><br>

仲量联行将定期修改本网站中包含的信息、服务和资源，并保留做此修改的权利，而无需另行通知过去、当前或者未来的访问者。
<br><br>
<b class="b">如何使用本网站</b><br>
倘若可以看到版权、商标或者其它所有权通告并且没有变化，您下载的内容只可用于非商业和个人目的。下载材料中的权利、所有权或者利益不会由于您的下载或者复制产生任何实质的转让。您同意，在没有仲量联行书面同意的情况下，不会以任何方式复制、修改、改变、展示、分发、出售、广播或者传送本网站上的任何材料。
<br><br>
<b class="b">不得用于非法或者禁止目的</b><br>
作为使用网站的条件之一，您不得将网站用于本使用条款或者任何适用法律禁止或者非法的目的。未经允许的意见提交
<br><br>
仲量联行不接受或者不考虑来自公众的任何创意、建议或者材料（“意见”），因此您不应向仲量联行提交任何意见。如果您不顾我们的善意要求坚持向我们发送意见，提交的意见将被认为是非保密的和非专利性的，并将立即成为仲量联行的财产。仲量联行现在和今后将独家拥有与此相关的一切权利、所有权和利益。仲量联行还可随意将此意见用于任何目的。
<br><br>
<b class="b">保密制度</b><br>
仲量联行对任何您提交到网站的个人信息的使用受网站保密制度的制约。
<br><br>
<b class="b">免责声明</b><br>
本网站建立在“现状（AS IS）”的基础之上。仲量联行明示不承担任何担保责任，包括商业适售性、特定目的之适用性及未侵害他人权利等担保。仲量联行对同以下相关或由此产生的任何类型的损失、伤害、索赔、债务或损害不承担任何责任：(A) 本网站及内容中的错误或遗漏，包括但不限于技术错误或排字错误；(B)通过本网站中的链接直接或间接访问的第三方网站或内容，包括但不限于其中的错误或遗漏；(C) 网站或者其中某部分内容无法使用；(D)您对本网站的使用或者；(E)您对同本网站连接的任何设备或软件的使用。
<br><br>
<b class="b">责任限制</b><br>
在一切情况下以及任何法律或者衡平理论下，无论是否存在民事侵权行为、合同、严格的责任，仲量联行对于使用在此包含信息所导致的直接、间接、特殊、偶然或者附带的损害不承担任何责任，包括但不限于：利润损失、商誉损失、数据丢失、工作停止、结果准确性、或者计算机故障造成的损害。
<br><br>
<b class="b">赔偿</b><br>
您同意保护、赔偿以及保障仲量联行免于承担由于您使用本网站导致和与之相关的包括律师费在内的任何索赔、损害、费用和开支。
<br><br>
<b class="b">版权通告</b><br>
除非另有规定，本网站所包含的图像、按钮及文字都是仲量联行及其分公司的专有财产。除个人使用外，在事先未经仲量联行的书面同意，这些项目不得以电子、机械、影印、录音或其它任何形式进行拷贝、分发、展示、复制或者传输。
<br><br>
<b class="b">商标</b><br>
本网站特色性的标识、品牌和其它商标以及服务标识（统称为“标识”）是仲量联行及其分公司的财产或其所获的授权。未经仲量联行书面同意，本网站上的任何内容都不应通过暗示、禁止反言或者其它方式解释为许可或授权使用本网站上显示的任何标识，或者任何第三方拥有本网站上显示的标识。
<br><br>
<b class="b">与第三方网站的链接</b><br>
为了方便用户，本网站可与第三方拥有并经营的其它网站进行链接，仲量联行不负责此类网站的维护。但是，即使这些第三方加盟仲量联行，仲量联行也无法管理这些链接网站，所有的网站都有各自的机密和数据收集方法，以及独立于仲量联行的法律政策。仲量联行对链接网站的内容不承担责任，并且对此类网站的内容或者材料的准确性不做任何陈述。访问这些第三方网站的风险完全由您自己承担。
<br><br>
<b class="b">关于前瞻性陈述的警告</b><br>
本网站会包含构成“前瞻性评论”的陈述、评判或推测如美国联邦证券法所定义。这些前瞻性陈述本身是通过推测或者基于当前可利用的信息、经营计划和对未来事件和趋势的预测，因此存在一定的风险和不确定性。实际结果和业绩会与仲量联行的过去经验和我们当前的期望或预测有实质性的差异。仲量联行无义务公开更新或者修改任何前瞻性陈述。
<br><br>
<b class="b">管辖权</b><br>
仲量联行在其位于美国伊利诺伊州芝加哥的办事处对本网站进行维护和运行。这些使用条款受美国伊利诺伊州法律的管辖并依据其进行司法解释。使用本网站表明您同意本使用条款导致的起诉由位于伊利诺伊州的法庭进行管辖。您对本网站持有的起诉理由或索赔必须在索赔或起诉理由产生后一年之内开始。如果这些使用条款中的任何部分被认为不合法、无效或者不可执行，那么该部分条款应被视为从本使用条款中分离并根据适用的法律进行解释。这一条款不会影响本使用条款中任何其它条款的合法有效性和执行力。仲量联行对于您或者他人违反这些使用条款未能采取措施并不构成对本使用条款的弃权，同时这次违约或者随后的违约将不会限制仲量联行的权利。
<br><br>
<b class="b">联系地址</b><br>
中国上海市静安区石门一路288号 兴业太古汇香港兴业中心一座22楼

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