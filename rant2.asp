
    <ul class="fang hidden">
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
    
    <li><a href="housedetails.asp?id=<%=rs("articleid")%>" class="fpic"><%if rs("gid")<>"" and rs("gid")<>0 then%><i class="hot"></i><%end if%><img src="<%=rs("defaultpicurl")%>" class="fangpic b_h"></a><div class="ninfo"><a  href="housedetails.asp?id=<%=rs("articleid")%>" class="ntit b_h ellipsis"><%=rs("title")%></a><span class="ntxt"><span class="ntxtl">区域<br>户型<br>豪宅类型<br></span><span class="ntxtr"><%=rs("bigclassname")%>，<%=rs("smallclassname")%><br>
<%
    mystr=split(rs("clxid"),",")
      for i=0 to ubound(mystr)
if i<4 and mystr(i)<>"" then
%>
<em class="trait"><%call Showtd(mystr(i),"lx")%></em>
<% end if
 next %><br><%=rs("clbname")%><br></span>
    <div class="c"></div>
    <div class="lcl">
    <%
        mystr=split(rs("ctdid"),",")
          for i=0 to ubound(mystr)
    if i<4 and mystr(i)<>"" then
    %>
    <i class="cl<%=i%> "><%call Showtd(mystr(i),"td")%></i>

    <% end if
     next %>
    </div>
    </span>
    </div><div class="nphone"><%if rs("jgzj")=0 then%><span class="jiage"><b class="b1">价格待定</b></span><%else%><span class="jiage">¥<%=rs("jgzj")%><b class="b1">万起</b><%end if%><!--<b class="b2">(每平)</b>--></span><div class="c hidden-xs"></div><a href="housedetails.asp?id=<%=rs("articleid")%>" class="nbtn">查看更多</a></div></li>
<%
rs.movenext
i=i+1
loop
end if
	rs.close
	set rs=nothing%>