      <aside>
          <div id="sidebar"  class="nav-collapse ">
              <!-- sidebar menu start-->
              <ul class="sidebar-menu">
                  <li class="nav0">
                      <a class="" href="index.asp">
                          <i class="icon-dashboard"></i>
                          <span>后台首页</span>
                      </a>
                  </li>
                  
                  <li class="sub-menu nav1">
                      <a href="javascript:;" class="">
                          <i class="icon-book"></i>
                          <span>网站设置</span>
                          <span class="arrow"></span>
                      </a>
                      <ul class="sub">
                          <li class="nav1-1"><a class="" href="sitebase.asp">基本信息管理</a></li>
                          <li class="nav1-2"><a class="" href="sitemanager.asp">管理员管理</a></li>
                          <li class="nav1-3"><a class="" href="sitespase.asp">空间使用信息</a></li>
                          <li class="nav1-4"><a class="" href="lyb.asp">网站报名管理</a></li>
                          <li class="nav1-5"><a class="" href="tdk_list.asp">网站SEO设置</a></li>
                          <!-- <li class="nav1-4"><a class="" href="siteinfo.asp">服务器情况</a></li>-->
                      </ul>
                  </li>
<!--                  <li class="sub-menu nav2">
                          <a href="javascript:;" class="">
                              <i class="icon-cogs"></i>
                              <span>关于我们</span>
                              <span class="arrow"></span>
                          </a>
                          <ul class="sub">
                              <li class="nav2-1"><a class="" href="single.asp?nav=1">企业介绍</a></li>
                              <li class="nav2-2"><a class="" href="single.asp?nav=2">国际团队</a></li>
                              <li class="nav2-3"><a class="" href="class_links.asp">友情链接</a></li>
                          </ul>
                      </li>-->
                   <li class="sub-menu nav3">
                      <a href="javascript:;" class="">
                          <i class="icon-tasks"></i>
                          <span>房源管理</span>
                          <span class="arrow"></span>
                      </a>
                      <ul class="sub">
                          <li class="nav3-1"><a class="" href="Pro_List.Asp">房源列表</a></li>
                          <li class="nav3-2"><a class="" href="Pro_Add.Asp">发布房源</a></li>
                          <li class="nav3-3"><a class="" href="Pro_Class.Asp">地域分类</a></li>
                          <li class="nav3-4"><a class="" href="class_lx.asp">房源面积</a></li>
                          <li class="nav3-5"><a class="" href="class_lb.asp">房源类别</a></li>
                          <li class="nav3-6"><a class="" href="class_fj.asp">价格区间</a></li>
                          <li class="nav3-7"><a class="" href="class_td.asp">房源特点</a></li>
                          <li class="nav3-8"><a class="" href="class_rq.asp">房源厅室</a></li>
<!--<li class="nav3-8"><a class="" href="class_kfs.asp">开发商</a></li>-->
                      </ul>
                  </li>
  <li class="sub-menu nav4">
                      <a href="javascript:;" class="">
                          <i class="icon-tasks"></i>
                          <span>最新资讯</span>
                          <span class="arrow"></span>
                      </a>
                      <ul class="sub">
                    <!--  <li class="nav4-5"><a class="" href="class_prod.asp">新闻类别管理</a></li>-->
<!--                          <li class="nav4-1"><a class="" href="Prod_list.Asp?Prod_Type=11">展会资讯列表</a></li>
                          <li class="nav4-2"><a class="" href="Prod_add.asp?Prod_Type=11">增添展会资讯</a></li>
                          <li class="nav4-3"><a class="" href="Prod_list.Asp?Prod_Type=12">行业动态列表</a></li>
                          <li class="nav4-4"><a class="" href="Prod_add.asp?Prod_Type=12">增添行业动态</a></li>-->
                          <%set rs=Server.CreateObject("ADODB.Recordset")
rs.Open "select * from [class_prod] where cid<>8 order by oid desc,cid desc",conn,1,1
if rs.bof and rs.eof then 
else 

i=0
do while not rs.eof

%>
<li class="nav4-<%=rs("cid")%>"><a class="" href="Prod_list.Asp?Prod_Type=<%=rs("cid")%>"><%=rs("cname")%>列表</a></li>
                          <li class="nav4-<%=rs("cid")%>a"><a class="" href="Prod_add.asp?Prod_Type=<%=rs("cid")%>">增添<%=rs("cname")%></a></li>

<%
rs.movenext
	i=i+1
    loop
end if
rs.close
	set rs=nothing%>
                      </ul>
                  </li>
 <li class="sub-menu nav5">
                          <a href="javascript:;" class="">
                              <i class="icon-tasks"></i>
                              <span>销售团队管理</span>
                              <span class="arrow"></span>
                          </a>
                          <ul class="sub">
                              <li class="nav5-1"><a class="" href="team_list.Asp?classid=1">销售团队列表</a></li>
                              <li class="nav5-2"><a class="" href="team_add.asp?classid=1">添加团队成员</a></li>
                          </ul>
                      </li>
                  <li class="sub-menu nav6">
                                        <a href="javascript:;" class="">
                                            <i class="icon-tasks"></i>
                                            <span>广告链接</span>
                                            <span class="arrow"></span>
                                        </a>
                                        <ul class="sub">
                                            <li class="nav6-1"><a class="" href="links.asp?c=1">首页广告管理</a></li>
                                            <li class="nav6-2"><a class="" href="links.asp?c=2">新闻广告管理</a></li>
                                        </ul>
                                    </li>
                  <li>
                      <a class="" href="LogInout.Asp">
                          <i class="icon-user"></i>
                          <span>安全退出</span>
                      </a>
                  </li>
              </ul>
              <!-- sidebar menu end-->
          </div>
      </aside>