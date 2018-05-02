
<!--#include file="inc/xgheader.asp"-->
      <!--sidebar start-->
<!--#include file="inc/xgmenu.asp"-->
      <!--sidebar end-->
      <!--main content start-->
      <section id="main-content">
          <section class="wrapper">
          <%
dim rs
dim sql
dim count

set rs=server.createobject("adodb.recordset")
sql = "select * from [Table_ProSmallclass] order by orderid desc,SmallClassID desc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>;<%= trim(rs("SmallClassId"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid.split(";")[0];
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }  
	
	</script>
              <!-- page start -->

                            <!--成功失败提示-->
<div class="alert" id="alert" style="display:none;">
                                  <button class="close close-sm" type="button">
                                      <i class="icon-remove"></i>
                                  </button>
                                  <strong>操作成功！</strong> <span>栏目内容修改成功。</span>
                              </div>
 
                          <!----> 
              <div class="row ">
                  <div class="col-lg-12">
                      <section class="panel pb50">
                          <header class="panel-heading mb50">
                              信息修改
                          </header>
                      <form name="myform" id="myform">    
                      <%
Dim PID
ID=Trim(Request.QueryString("ID"))
classid=Trim(Request.QueryString("Prod_Type"))
set oRs=Server.CreateObject("ADODB.Recordset")
oRs.Open "select * from [team] where id="&ID,conn,1,3
if oRs.Eof then
'无数据
End if
%>
<!--国家分类开始-->

  <!--国家分类结束-->
    <div class="form-group row">
     <label for="cname" class="col-sm-2 control-label text-right xglab">中文名</label>
     <div class="col-sm-7">
       <input type="text" name="cname" class="form-control" id="cname" value="<%=oRs("cname")%>" >
     </div>
   </div>
       <div class="form-group row">
       <label for="ename" class="col-sm-2 control-label text-right xglab">英文名</label>
       <div class="col-sm-7">
         <input type="text" name="ename" class="form-control" id="ename"  value="<%=oRs("ename")%>">
       </div>
     </div>
           <div class="form-group row">
           <label for="position" class="col-sm-2 control-label text-right xglab">职位</label>
           <div class="col-sm-7">
             <input type="text" name="position" class="form-control" id="position"  value="<%=oRs("position")%>">
           </div>
         </div>
         <div class="form-group row">
                   <label for="tel" class="col-sm-2 control-label text-right xglab">电话</label>
                   <div class="col-sm-7">
                     <input type="text" name="tel" class="form-control" id="tel"  value="<%=oRs("tel")%>">
                   </div>
                 </div>
               <div class="form-group row">
                      <label for="mail" class="col-sm-2 control-label text-right xglab">邮箱</label>
                       <div class="col-sm-7">
                      <input type="text" name="mail" class="form-control" id="mail"  value="<%=oRs("mail")%>">
                      </div>
                 </div>
                 <div class="form-group row">
                                   <label for="summary" class="col-sm-2 control-label text-right xglab">摘要</label>
                                   <div class="col-sm-7">
                                     <input type="text" name="summary" class="form-control" id="summary"  value="<%=oRs("summary")%>">
                                   </div>
                                 </div>
       <div class="form-group row">
                                         <label for="info" class="col-sm-2 control-label text-right xglab">介绍</label>
                                         <div class="col-sm-7">
                                           <input type="text" name="info" class="form-control" id="info"  value="<%=oRs("info")%>">
                                         </div>
                                       </div>


       <div class="form-group row">
     <label for="pic" class="col-sm-2 control-label text-right xglab">照片</label>
     <div class="col-sm-7 row">
       <div class="col-sm-8"><input name="pic" type="text" id="pic" class="form-control col-sm-6"  value="<%=oRs("pic")%>"></div>
         <div class="col-sm-3"><input type="button" name="Submit12" value="上传照片" class="form-control col-sm-2" onClick="window.open('Yr_UpFile.Asp?formname=myform&editname=pic&uppath=images/guanggao&filelx=jpg','','status=no,scrollbaRs=no,top=20,left=110,width=420,height=165')"></div>
     </div>
   </div>

         <div class="form-group row">
     <label for="uptime" class="col-sm-2 control-label text-right xglab">更新时间</label>
     <div class="col-sm-10 row">
     <div class="col-sm-4 input-group">
       <input type="text" name="uptime" class="form-control" id="uptime"  value="<%=oRs("uptime")%>" >
     </div>
     </div>
   </div>
     <div class="form-group row">
     <label for="oid" class="col-sm-2 control-label text-right xglab">排序ID</label>
     <div class="col-sm-7 row">
     <div class="col-sm-2">
       <input type="text" name="oid" class="form-control" id="oid"  value="<%=oRs("oid")%>" placeholder="数字越大排序越靠前" >
       </div>
     </div>
   </div>
<div class="row">
<div class="col-lg-offset-2 col-lg-7">
<button class="btn btn-danger save" type="button" style="margin:15px auto;">保存</button>
<input id="action" value="mod" type="hidden">
<input id="classid" value="<%=oRs("cid")%>" type="hidden">
<input id="id" value="<%=oRs("id")%>" type="hidden">
</div>
  </div>
</form>
<!--列表结束-->
                      </section>
                  </div>
              </div>

              <!--  page end  -->
          </section>
      </section>
      <!--main content end-->
  </section>

<!--#include file="inc/xgfooter.asp"-->
<!--#include file="inc/windows.asp"-->
  <script type="text/javascript">
//导航菜单

$(".nav5").addClass('active');
$(".nav5-2").addClass('active');


</script>
<script type="text/javascript" src="ajax/team.js"></script>
