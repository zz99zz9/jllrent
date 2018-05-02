<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/2.css"/>
<!--广告部份-->
<%cid=request.QueryString("cid")
if cid="" then cid=0
	  set rsnew=Server.CreateObject("ADODB.Recordset")
	  if cid<>0 then
	  call SafeRequest(cid)
	  sqlnew="select * from [prod] where classid="&cid
	  else
	  sqlnew="select * from [prod] where classid=11 or classid=12 "
	  end if

	  sqlnew=sqlnew+" order by oid desc,prod_id desc"

rsnew.Open sqlnew,conn,1,1%>
<div class="navwz container">
<a href="#">JLL</a> &gt; <a href="#">最新资讯</a>
</div>
<!--广告部份-->
<div class="ed2">
<div class="container">
<div id="carousel-example-generic" class="carousel slide" data-ride="carousel">

        <ol class="carousel-indicators">
            <li data-target="#carousel-example-generic" data-slide-to="0" class="active"></li>
            <li data-target="#carousel-example-generic" data-slide-to="1"></li>
            <li data-target="#carousel-example-generic" data-slide-to="2"></li>
        </ol>


        <div class="carousel-inner" role="listbox">



<%Set prs= Server.CreateObject("ADODB.Recordset")
                                psql="select top 3 * From [links] where wid=2 order by oid desc,ID desc"
                                prs.open psql,conn,1,1
                                i=0%>
                                <%do while not prs.eof%>
                          <div class="item <%if i=0 then%>active<%end if%>"  onclick="javascript:window.open('<%=prs("url")%>');" class="hand" >
                         <img src="<%=prs("file")%>"  alt="...">
                               <div class="info">
                                            <h1><%=prs("stit")%></h1>
                                              <p class="tra"><%=prs("stxt")%></p>
                      </div>
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
    </div>
</div>
<!--part1-->
<div class="part1">
<div class="container">
    <ul class="nav2"><a class="hand <%if cid=0 then%>on<%end if%>" href="2.asp">全部</a><a class=" hand <%if cid=11 then%>on<%end if%>" href="2.asp?cid=11">豪宅研究报告</a><a class="hand <%if cid=12 then%>on<%end if%>" href="2.asp?cid=12">最新活动</a><!--<a class="hand <%if cid=15 then%>on<%end if%>" href="2.asp?cid=15">生活品鉴</a>--></ul>
    <ul class="news">
     <%
            if rsnew.bof and rsnew.eof then
            response.write("<br><br><br>&nbsp;&nbsp;内容整理中……")
            else
            rsnew.PageSize=15 '设置页码
            pagecount=rsnew.PageCount '获取总页码
            page=int(request("page")) '接收页码
            if page<=0 then page=1 '判断
            if request("page")="" then page=1
            rsnew.AbsolutePage=page '设置本页页码
            i=0
            do while not rsnew.eof and i<rsnew.PageSize
            %>
    <li><a href="newsdetail.asp?id=<%=rsnew("prod_id")%>"><img src="<%=rsnew("prod_pic")%>" class=" b_c tra"></a><div class="ninfo"><span class="ntime"><%=formatdatetime(rsnew("prod_date"),2)%></span><a class="ntit" href="newsdetail.asp?id=<%=rsnew("prod_id")%>"><%=rsnew("prod_name")%></a><span class="ntxt"><%=left(RemoveHTML(rsnew("prod_detail")),100)%>……</span></div><a href="newsdetail.asp?id=<%=rsnew("prod_id")%>" class="nbtn">查看更多</a></li>
              <%rsnew.movenext
            	i=i+1
                loop
            	end if
            	rsnew.close
            	set rsnew=nothing%>
 </ul>
</div>
</div>
<!--#include file="inc/footer.asp"-->